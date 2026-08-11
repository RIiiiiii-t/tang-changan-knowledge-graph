import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2] ?? "05-Excel数据/小型样例_第1批_文献片段与基础对象_v2.xlsx";
const outputDir = process.argv[3] ?? "05-Excel数据/平台导入包/唐长安小型样例_v1";
const generatedAt = new Date().toISOString();
const namespaceUrl = "6ba7b811-9dad-11d1-80b4-00c04fd430c8";

function uuidToBytes(uuid) {
  return Buffer.from(uuid.replaceAll("-", ""), "hex");
}

function uuidv5(name, namespace = namespaceUrl) {
  const hash = crypto.createHash("sha1").update(Buffer.concat([uuidToBytes(namespace), Buffer.from(name)])).digest();
  hash[6] = (hash[6] & 0x0f) | 0x50;
  hash[8] = (hash[8] & 0x3f) | 0x80;
  const hex = hash.subarray(0, 16).toString("hex");
  return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
}

const idFor = (kind, key) => uuidv5(`tang-changan:${kind}:${key}`);
const cleanHeader = value => String(value ?? "").replace(/\*$/, "").trim();
const splitKeys = value => String(value ?? "").split(/[,，;；\n]+/).map(x => x.trim()).filter(Boolean);
const aliases = value => splitKeys(value);
const compact = object => Object.fromEntries(Object.entries(object).filter(([,v]) => v !== null && v !== undefined && v !== ""));
const statusMap = value => ({"审核通过":"approved","已发布":"published","审核未通过":"rejected","退回修改":"rejected","存在争议":"disputed","待审核":"pending_review","等待确认":"pending_review"}[String(value ?? "").trim()] ?? "draft");
const evidenceVisibility = status => status === "published" ? "public" : "researcher";
const relationLabelMap = {INTERPRETS:"阐释",EXPLAINS:"解释",REFINES:"细化",SUPPORTS:"支持",CONTRADICTS:"冲突",AGREES_WITH:"支持另一结论",BASED_ON:"基于",PROPOSED_BY:"由…提出"};
function excelDateToIso(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return value.toISOString().slice(0,10);
  if (typeof value === "number") return new Date(Date.UTC(1899,11,30) + value * 86400000).toISOString().slice(0,10);
  const text = String(value).trim();
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? text : date.toISOString().slice(0,10);
}
function pageFields(value) {
  if (value === null || value === undefined || value === "") return {};
  const pageLabel = String(value).trim();
  const journalPage = pageLabel.match(/刊页\s*(\d+)/);
  if (journalPage) return {page_label:pageLabel,page_start:Number(journalPage[1]),page_end:Number(journalPage[1])};
  const range = pageLabel.match(/^(\d+)\s*[—–-]\s*(\d+)$/);
  if (range) return {page_label:pageLabel,page_start:Number(range[1]),page_end:Number(range[2])};
  const single = pageLabel.match(/^(\d+)$/);
  if (single) return {page_label:pageLabel,page_start:Number(single[1]),page_end:Number(single[1])};
  return {page_label:pageLabel};
}

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
function rows(sheetName) {
  const values = wb.worksheets.getItem(sheetName).getUsedRange(true).values;
  const headers = values[2].map(cleanHeader);
  return values.slice(3).filter(row => row.some(v => v !== null && v !== "")).map((row,index) => {
    const record = {__sheet:sheetName,__row:index+4};
    headers.forEach((header,i) => { if (header) record[header] = row[i]; });
    return record;
  });
}

const issues = [];
const error = (code,message,context={}) => issues.push({severity:"ERROR",code,message,...context});
const warn = (code,message,context={}) => issues.push({severity:"WARNING",code,message,...context});

const sourceRows = rows("来源表");
const chunkRows = rows("史料片段");
const sourceKeys = new Set(sourceRows.map(r => r.source_key));
const chunkKeys = new Set(chunkRows.map(r => r.chunk_key));

const sources = sourceRows.map(r => {
  const status = statusMap(r.review_status);
  const qualityConfidence = ({A:0.95,B:0.85,C:0.70,D:0.50})[r.quality_level] ?? 0.70;
  return {
    id:idFor("source",r.source_key),
    title:r.title,
    source_type:r.source_type,
    identifier:r.source_key,
    payload:compact({project_key:r.source_key,author:r.author,edition:r.edition,publisher:r.publisher,publication_year:r.publication_year,volume:r.volume,page:r.page,external_identifier:r.identifier,url_or_file:r.url_or_file,quality_level:r.quality_level,notes:r.notes}),
    status,
    visibility:evidenceVisibility(status),
    confidence:qualityConfidence,
    data_classification:"research_pilot"
  };
});

const sourceIdByKey = new Map(sources.map(x => [x.identifier,x.id]));
const sourceChunks = chunkRows.map(r => {
  if (!sourceKeys.has(r.source_key)) error("MISSING_SOURCE",`片段 ${r.chunk_key} 引用不存在的来源 ${r.source_key}`,{sheet:r.__sheet,row:r.__row,key:r.chunk_key});
  const status = statusMap(r.review_status);
  return {
    id:idFor("chunk",r.chunk_key),key:r.chunk_key,source_id:sourceIdByKey.get(r.source_key),original_text:r.content,
    payload:compact({chapter:r.chapter,volume:r.volume,page:r.page,...pageFields(r.page),related_entity_keys:splitKeys(r.related_entity_keys),evidence_type:r.evidence_type,note:r.note,correction_status:status === "approved" ? "human_reviewed" : "unreviewed"}),
    status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7),created_at:null,updated_at:null,data_classification:"research_pilot"
  };
});
const chunkIdByKey = new Map(sourceChunks.map(x => [x.key,x.id]));

function evidenceIds(record) {
  const keys = splitKeys(record.chunk_key);
  for (const key of keys) if (!chunkKeys.has(key)) error("MISSING_CHUNK",`${record.__sheet}第${record.__row}行引用不存在的片段 ${key}`,{sheet:record.__sheet,row:record.__row,key:key});
  return keys.map(k => chunkIdByKey.get(k)).filter(Boolean);
}

const entities = [];
const entityKeySet = new Set();
function addEntity(entity, origin) {
  if (!entity.key) return error("MISSING_ENTITY_KEY",`${origin.__sheet}第${origin.__row}行缺少实体编号`,{sheet:origin.__sheet,row:origin.__row});
  if (entityKeySet.has(entity.key)) return error("DUPLICATE_ENTITY_KEY",`实体编号重复：${entity.key}`,{sheet:origin.__sheet,row:origin.__row,key:entity.key});
  entityKeySet.add(entity.key);
  entities.push({...entity,id:idFor("entity",entity.key),normalized_name:entity.name,payload:{...entity,origin_sheet:origin.__sheet,origin_row:origin.__row},created_at:null,updated_at:null,data_classification:"research_pilot"});
}

for (const r of rows("实体表")) {
  const status=statusMap(r.review_status);
  addEntity({key:r.entity_key,label:r.entity_type,name:r.name,aliases:aliases(r.aliases),properties:compact({knowledge_layer:"FACT",name_traditional:r.name_traditional,category:r.category,period_start:r.period_start,period_end:r.period_end,description:r.description,location_text:r.location_text,source_key:r.source_key,description_quote:r.description_quote}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7)},r);
}

for (const r of rows("功能分区规则")) {
  const status=statusMap(r.review_status);
  addEntity({key:r.rule_key,label:"FunctionalZoningRule",name:r.rule_name,aliases:[],properties:compact({knowledge_layer:"INTERPRETATION",rule_type:r.rule_type,scope:r.scope,condition_text:r.condition_text,conclusion_text:r.conclusion_text,related_fang_keys:splitKeys(r.related_fang_keys),related_road_keys:splitKeys(r.related_road_keys),related_canal_keys:splitKeys(r.related_canal_keys),related_market_keys:splitKeys(r.related_market_keys),related_palace_keys:splitKeys(r.related_palace_keys),positive_examples:r.positive_examples,exceptions:r.exceptions,inference_steps:r.inference_steps,period_start:r.period_start,period_end:r.period_end,evidence_type:r.evidence_type,note:r.note}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7)},r);
}

for (const r of rows("坊内平面规则")) {
  const status=statusMap(r.review_status);
  addEntity({key:r.layout_rule_key,label:"LayoutRule",name:r.rule_name,aliases:[],properties:compact({knowledge_layer:"INTERPRETATION",rule_type:r.rule_type,description:r.description,applies_to_keys:splitKeys(r.applies_to_keys),period_start:r.period_start,period_end:r.period_end,evidence_type:r.evidence_type,note:r.note}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7)},r);
}

for (const r of rows("尺度关系表")) {
  const status=statusMap(r.review_status);
  addEntity({key:r.scale_key,label:"ScaleComparison",name:`${r.subject_a_name}与${r.subject_b_name}${r.metric}比较`,aliases:[],properties:compact({knowledge_layer:"FACT",subject_a_key:r.subject_a_key,subject_b_key:r.subject_b_key,metric:r.metric,a_original_value:r.a_original_value,a_original_unit:r.a_original_unit,a_normalized_value:r.a_normalized_value,b_original_value:r.b_original_value,b_original_unit:r.b_original_unit,b_normalized_value:r.b_normalized_value,normalized_unit:r.normalized_unit,ratio_a_to_b:r.ratio_a_to_b,calculation_note:r.calculation_note,comparability:r.comparability,measurement_method:r.measurement_method,uncertainty_note:r.uncertainty_note,period_start:r.period_start,period_end:r.period_end,evidence_type:r.evidence_type,note:r.note}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7)},r);
}

const claimRows=rows("研究结论表");
for (const r of claimRows) {
  const status=statusMap(r.review_status);
  addEntity({key:r.claim_key,label:"InterpretiveClaim",name:r.claim_title,aliases:[],properties:compact({knowledge_layer:"INTERPRETATION",claim_text:r.claim_text,popular_text:r.popular_text,claim_type:r.claim_type,subject_keys:splitKeys(r.subject_keys),subject_relation_type:r.subject_relation_type,proposed_by:r.proposed_by,basis_text:r.basis_text,reasoning_chain:r.reasoning_chain,scope:r.scope,limitations:r.limitations,claim_status:r.claim_status,note:r.note}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence ?? 0.7)},r);
}

const popularRows=rows("科普内容表");
const popularLinks=rows("科普依据关联");
for (const r of popularRows) {
  const directChunkLinks=popularLinks.filter(x=>x.content_key===r.content_key && x.evidence_object_type==="CHUNK");
  const directChunkKeys=directChunkLinks.map(x=>x.evidence_key);
  const directEvidenceLinks=directChunkLinks.map(x=>compact({link_key:x.link_key,evidence_key:x.evidence_key,evidence_id:chunkIdByKey.get(x.evidence_key),support_role:x.support_role,source_layer:x.source_layer,display_order:x.display_order,public_explanation:x.public_explanation,is_visible:String(x.is_visible??"").trim()!=="否",review_status:statusMap(x.review_status),note:x.note}));
  const fake={...r,chunk_key:directChunkKeys.join(",")};
  const status=statusMap(r.science_review_status);
  addEntity({key:r.content_key,label:"PopularContent",name:r.title,aliases:[],properties:compact({knowledge_layer:"POPULAR",question:r.question,short_answer:r.short_answer,full_text:r.full_text,content_type:r.content_type,target_audience:r.target_audience,difficulty_level:r.difficulty_level,media_form:r.media_form,narrative_perspective:r.narrative_perspective,interaction_prompt:r.interaction_prompt,content_version:r.version,generated_by:r.generated_by,science_review_status:r.science_review_status,editorial_review_status:r.editorial_review_status,reviewer:r.reviewer,created_date:excelDateToIso(r.created_date),updated_date:excelDateToIso(r.updated_date),direct_evidence_links:directEvidenceLinks,note:r.note}),source_ids:evidenceIds(fake),status,visibility:"researcher",confidence:0.7},r);
}

const entityIdByKey=new Map(entities.map(x=>[x.key,x.id]));
const entityByKey=new Map(entities.map(x=>[x.key,x]));
const relations=[];
const relationKeySet=new Set();
function addRelation(relation,origin) {
  if (relationKeySet.has(relation.key)) return error("DUPLICATE_RELATION_KEY",`关系编号重复：${relation.key}`,{sheet:origin.__sheet,row:origin.__row,key:relation.key});
  relationKeySet.add(relation.key);
  if(!entityIdByKey.has(relation.source_key)) error("MISSING_SOURCE_ENTITY",`关系 ${relation.key} 的起点不存在：${relation.source_key}`,{sheet:origin.__sheet,row:origin.__row,key:relation.key});
  if(!entityIdByKey.has(relation.target_key)) error("MISSING_TARGET_ENTITY",`关系 ${relation.key} 的终点不存在：${relation.target_key}`,{sheet:origin.__sheet,row:origin.__row,key:relation.key});
  if(!relation.source_ids?.length) error("RELATION_WITHOUT_EVIDENCE",`关系 ${relation.key} 没有来源片段`,{sheet:origin.__sheet,row:origin.__row,key:relation.key});
  relations.push({id:idFor("relation",relation.key),key:relation.key,source_id:entityIdByKey.get(relation.source_key),target_id:entityIdByKey.get(relation.target_key),type:relation.type,label:relation.label,properties:{...relation.properties,source_key:relation.source_key,target_key:relation.target_key},source_ids:relation.source_ids,status:relation.status,visibility:relation.visibility,confidence:relation.confidence,payload:{...relation,origin_sheet:origin.__sheet,origin_row:origin.__row},created_at:null,updated_at:null,data_classification:"research_pilot"});
}

for(const r of rows("关系表")){
 const status=statusMap(r.review_status);
 addRelation({key:r.relation_key,source_key:r.source_key,target_key:r.target_key,type:r.relation_type,label:r.relation_label,properties:compact({knowledge_layer:r.evidence_type==="推导结论"?"INTERPRETATION":"FACT",period_start:r.period_start,period_end:r.period_end,evidence_quote:r.evidence_quote,evidence_type:r.evidence_type,note:r.note}),source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence??0.7)},r);
}

for(const r of claimRows){
 const targets=splitKeys(r.subject_keys); const status=statusMap(r.review_status);
 targets.forEach((target,index)=>addRelation({key:`TC-REL-${r.claim_key.replace("TC-CLAIM-", "CLAIM-")}-${String(index+1).padStart(2,"0")}`,source_key:r.claim_key,target_key:target,type:r.subject_relation_type||"INTERPRETS",label:relationLabelMap[r.subject_relation_type]||"解释",properties:{knowledge_layer:"INTERPRETATION",proposed_by:r.proposed_by,scope:r.scope,limitations:r.limitations},source_ids:evidenceIds(r),status,visibility:evidenceVisibility(status),confidence:Number(r.confidence??0.7)},r));
}

for(const r of popularLinks){
 if(r.evidence_object_type==="CHUNK" || r.evidence_object_type==="SOURCE" || r.evidence_object_type==="RELATION") continue;
 const target=entityByKey.get(r.evidence_key);
 const content=entityByKey.get(r.content_key);
 const status=statusMap(r.review_status);
 addRelation({key:r.link_key,source_key:r.content_key,target_key:r.evidence_key,type:"SUPPORTED_BY",label:"依据",properties:compact({knowledge_layer:"POPULAR",evidence_object_type:r.evidence_object_type,support_role:r.support_role,source_layer:r.source_layer,display_order:r.display_order,public_explanation:r.public_explanation,is_visible:r.is_visible,note:r.note}),source_ids:[...(content?.source_ids??[]),...(target?.source_ids??[])].filter((x,i,a)=>a.indexOf(x)===i),status,visibility:"researcher",confidence:0.7},r);
}

for(const source of sources) if(!source.title || !source.identifier) error("INVALID_SOURCE","来源缺少题名或编号",{key:source.identifier});
for(const chunk of sourceChunks) if(!chunk.original_text) error("EMPTY_CHUNK",`片段 ${chunk.key} 内容为空`,{key:chunk.key});
for(const entity of entities) if(!entity.source_ids.length) warn("ENTITY_WITHOUT_DIRECT_EVIDENCE",`实体 ${entity.key} 没有直接片段证据`,{key:entity.key});
const chunkById=new Map(sourceChunks.map(x=>[x.id,x]));
for(const object of [...entities,...relations]) if(object.status==="approved") {
  const unapproved=(object.source_ids??[]).map(id=>chunkById.get(id)).filter(chunk=>chunk && !["approved","published"].includes(chunk.status));
  if(unapproved.length) error("APPROVED_OBJECT_WITH_UNAPPROVED_EVIDENCE",`审核通过对象 ${object.key} 引用了未审核片段：${unapproved.map(x=>x.key).join(",")}`,{key:object.key});
}

const errors=issues.filter(x=>x.severity==="ERROR");
const warnings=issues.filter(x=>x.severity==="WARNING");
const packageData={contract_version:"TC-PLATFORM-CONTRACT-1.0",schema_version:"TC-SCHEMA-1.2",dataset_key:"tang_changan_pilot_v1",generated_at:generatedAt,source_workbook:path.basename(inputPath),sources,source_chunks:sourceChunks,entities,relations,validation_summary:{errors:errors.length,warnings:warnings.length,source_count:sources.length,source_chunk_count:sourceChunks.length,entity_count:entities.length,relation_count:relations.length,status_counts:[...entities,...relations].reduce((a,x)=>(a[x.status]=(a[x.status]??0)+1,a),{})}};

const lines=["# 唐长安小型样例平台导入校验报告 v1","",`生成时间：${generatedAt}`,`输入文件：${inputPath}`,`契约版本：${packageData.contract_version}`,"","## 结果", "",errors.length?`**未通过：发现 ${errors.length} 个错误。不得导入平台。`:`**通过结构校验：未发现阻断导入的错误。**`,"",`- 来源：${sources.length}`,`- 来源片段：${sourceChunks.length}`,`- 实体：${entities.length}`,`- 关系：${relations.length}`,`- 警告：${warnings.length}`,"","## 状态说明","",`本包包含 ${packageData.validation_summary.status_counts.pending_review??0} 条待审核数据、${packageData.validation_summary.status_counts.approved??0} 条审核通过数据。转换程序未把任何记录自动改为 published。`,`科普样例 TC-POPULAR-0001 保持 pending_review。`,"","## 错误","",...(errors.length?errors.map((x,i)=>`${i+1}. [${x.code}] ${x.message}`):["无。"]),"","## 警告","",...(warnings.length?warnings.map((x,i)=>`${i+1}. [${x.code}] ${x.message}`):["无。"]),"","## 人工复核清单","","- [ ] 数量是否符合当前小型样例范围","- [ ] TC-*编号是否完整且无重复","- [ ] 每条事实和关系是否可以回溯到来源片段","- [ ] 研究解释是否显示作者、范围和限制","- [ ] 已审核与待审核状态是否正确","- [ ] TC-POPULAR-0001是否保持未发布","- [ ] 重复运行程序所得UUID是否完全一致","","## 结论","",errors.length?"修复错误后重新生成。":"可以进入人工复核；人工确认后才可尝试导入测试数据库。",""];

await fs.mkdir(outputDir,{recursive:true});
await fs.writeFile(path.join(outputDir,"唐长安小型样例_平台导入包_v1.json"),JSON.stringify(packageData,null,2),"utf8");
await fs.writeFile(path.join(outputDir,"唐长安小型样例_导入校验报告_v1.md"),lines.join("\n"),"utf8");
await fs.writeFile(path.join(outputDir,"唐长安小型样例_校验问题_v1.json"),JSON.stringify(issues,null,2),"utf8");
console.log(JSON.stringify(packageData.validation_summary,null,2));
if(errors.length) process.exitCode=2;
