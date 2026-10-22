// Banco do painel na Vercel.
//
// Guarda ajustes, fechamentos e vistos num Redis (Upstash, via Vercel
// Marketplace) e serve o painel em painel/index.html. Sem dependências:
// fala com o Redis pela API REST.
//
// Variáveis de ambiente (Vercel → Settings → Environment Variables):
//   KV_REST_API_URL / KV_REST_API_TOKEN        (criadas pelo Upstash Redis)
//   ou UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN
//   PAINEL_SENHA                                (opcional) senha que o painel pede
//
// Rotas:
//   GET    /api/dados                 → {ajustes:{id:doc}, fechamentos:{}, vistos:{}, ver}
//   GET    /api/dados?ver=N           → {same:true} se nada mudou desde a versão N
//   PUT    /api/dados?c=<col>&id=<id> → grava o documento (corpo JSON)
//   DELETE /api/dados?c=<col>&id=<id> → apaga o documento
//
// Na primeira leitura com o banco vazio, importa ajustes/seed_painel.json
// (o que já tinha sido gravado no painel do Claude).

const COLS = ["ajustes", "fechamentos", "vistos"];
const KEY = c => `painel:${c}`;

function redisCfg(){
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if(!url || !token) return null;
  return { url: url.replace(/\/$/, ""), token };
}

async function redis(cfg, cmd){
  const r = await fetch(cfg.url, {
    method: "POST",
    headers: { Authorization: `Bearer ${cfg.token}`, "Content-Type": "application/json" },
    body: JSON.stringify(cmd),
  });
  const j = await r.json();
  if(!r.ok || j.error) throw new Error(`redis ${cmd[0]}: ${j.error || r.status}`);
  return j.result;
}

async function redisPipeline(cfg, cmds){
  for(let i = 0; i < cmds.length; i += 100){
    const lote = cmds.slice(i, i + 100);
    const r = await fetch(`${cfg.url}/pipeline`, {
      method: "POST",
      headers: { Authorization: `Bearer ${cfg.token}`, "Content-Type": "application/json" },
      body: JSON.stringify(lote),
    });
    const j = await r.json();
    if(!r.ok) throw new Error(`redis pipeline: ${r.status}`);
    const erro = (Array.isArray(j) ? j : []).find(x => x && x.error);
    if(erro) throw new Error(`redis pipeline: ${erro.error}`);
  }
}

function hashParaObjeto(flat){
  const o = {};
  if(!Array.isArray(flat)) return o;
  for(let i = 0; i + 1 < flat.length; i += 2){
    try{ o[flat[i]] = JSON.parse(flat[i + 1]); }catch(e){ /* ignora valor corrompido */ }
  }
  return o;
}

async function lerTudo(cfg){
  const out = {};
  for(const c of COLS) out[c] = hashParaObjeto(await redis(cfg, ["HGETALL", KEY(c)]));
  return out;
}

async function semearSePreciso(cfg){
  const feito = await redis(cfg, ["GET", "painel:seeded"]);
  if(feito) return false;
  let seed;
  try{ seed = require("../ajustes/seed_painel.json"); }catch(e){ seed = null; }
  const cmds = [];
  if(seed){
    for(const c of COLS){
      for(const [id, doc] of Object.entries(seed[c] || {})) cmds.push(["HSET", KEY(c), id, JSON.stringify(doc)]);
    }
  }
  cmds.push(["SET", "painel:seeded", new Date().toISOString()]);
  await redisPipeline(cfg, cmds);
  return true;
}

function corpo(req){
  if(req.body === undefined || req.body === null) return null;
  if(typeof req.body === "string"){ try{ return JSON.parse(req.body); }catch(e){ return null; } }
  if(Buffer.isBuffer(req.body)){ try{ return JSON.parse(req.body.toString("utf8")); }catch(e){ return null; } }
  return req.body;
}

function idValido(id){ return typeof id === "string" && id.length > 0 && id.length <= 300 && !/[\s\/]/.test(id); }

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("Content-Type", "application/json; charset=utf-8");

  const cfg = redisCfg();
  if(!cfg){
    res.statusCode = 503;
    return res.end(JSON.stringify({ erro: "Falta o banco: crie um Upstash Redis em Storage e conecte ao projeto (variáveis KV_REST_API_URL e KV_REST_API_TOKEN)." }));
  }
  // Senha é opcional: só é exigida se PAINEL_SENHA estiver definida na Vercel.
  const senha = process.env.PAINEL_SENHA;
  if(senha && req.headers["x-senha"] !== senha){
    res.statusCode = 401;
    return res.end(JSON.stringify({ erro: "senha" }));
  }

  const q = req.query || {};
  try{
    if(req.method === "GET"){
      await semearSePreciso(cfg);
      // ?ver=N: o painel pergunta se mudou algo desde a versão N; se não, resposta curta
      const ver = Number(await redis(cfg, ["GET", "painel:ver"])) || 0;
      if(q.ver !== undefined && Number(q.ver) === ver) return res.end(JSON.stringify({ same: true, ver }));
      const dados = await lerTudo(cfg);
      dados.pode_escrever = true;
      dados.ver = ver;
      return res.end(JSON.stringify(dados));
    }
    const c = q.c, id = q.id;
    if(!COLS.includes(c) || !idValido(id)){
      res.statusCode = 400;
      return res.end(JSON.stringify({ erro: "coleção ou id inválido" }));
    }
    if(req.method === "PUT"){
      const doc = corpo(req);
      if(!doc || typeof doc !== "object" || Array.isArray(doc)){
        res.statusCode = 400;
        return res.end(JSON.stringify({ erro: "corpo precisa ser um objeto JSON" }));
      }
      await redis(cfg, ["HSET", KEY(c), id, JSON.stringify(doc)]);
      const ver = await redis(cfg, ["INCR", "painel:ver"]);
      return res.end(JSON.stringify({ ok: true, ver }));
    }
    if(req.method === "DELETE"){
      await redis(cfg, ["HDEL", KEY(c), id]);
      const ver = await redis(cfg, ["INCR", "painel:ver"]);
      return res.end(JSON.stringify({ ok: true, ver }));
    }
    res.statusCode = 405;
    return res.end(JSON.stringify({ erro: "método" }));
  }catch(e){
    res.statusCode = 500;
    return res.end(JSON.stringify({ erro: String(e && e.message || e) }));
  }
};
