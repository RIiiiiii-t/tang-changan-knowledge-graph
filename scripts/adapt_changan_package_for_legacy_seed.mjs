import fs from "node:fs/promises";
import path from "node:path";

const inputPath = process.argv[2] ?? "05-Excel数据/平台导入包/唐长安小型样例_v1/唐长安小型样例_平台导入包_v1.json";
const outputPath = process.argv[3] ?? "05-Excel数据/平台导入包/唐长安小型样例_v1/唐长安小型样例_旧平台种子兼容包_v1.json";

const source = JSON.parse(await fs.readFile(inputPath, "utf8"));

function exposePayloadFields(record) {
  return {
    ...(record.payload ?? {}),
    ...record,
  };
}

const adapted = {
  ...source,
  compatibility_target: "legacy_seed_demo_as_params",
  compatibility_note: "将sources和source_chunks的payload字段同时暴露到记录顶层，供旧平台seed_demo读取；规范payload结构保持不变。",
  sources: source.sources.map(exposePayloadFields),
  source_chunks: source.source_chunks.map(exposePayloadFields),
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, JSON.stringify(adapted, null, 2), "utf8");

console.log(JSON.stringify({
  output: outputPath,
  sources: adapted.sources.length,
  source_chunks: adapted.source_chunks.length,
  entities: adapted.entities.length,
  relations: adapted.relations.length,
  sample_chunk: adapted.source_chunks.find(item => item.key === "TC-CHUNK-1102"),
}, null, 2));
