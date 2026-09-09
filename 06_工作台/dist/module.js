const modules = {
  a: { letter: 'A', name: 'ASIN 周环比监控', summary: '以全部转向器 ASIN 为范围，比较上周与上上周，定位销量、销售额或 TACOS 异常。', kpiLabel: '预警 ASIN', scope: 'VC1 转向器 ASIN · W35 vs W34', steps: [['刷新数据', '同步财务销量、销售额与广告活动数据。'], ['计算周环比', '比较单量、销售额、曝光、点击、CTR、CVR、ACOS、ROAS 与 TACOS。'], ['形成预警', '任一阈值命中后带出关联活动与广告指标。']], files: [['最终预警明细', '../../04_分析结果/模块A_最终预警明细表_2026-W35_vs_2026-W34.xlsx'], ['周环比监控表', '../../04_分析结果/模块A_ASIN周环比监控表_2026-W35_vs_2026-W34.csv']] },
  b: { letter: 'B', name: '掉量原因归因', summary: '以掉量起点前 14 天至后 7 天为窗口，按直接程度定位经营与广告异常。', kpiLabel: '归因结论', scope: '预警 / 掉量 ASIN · 日粒度 21 天窗口', steps: [['识别掉量起点', '连续 3 天低于此前 14 天日均销量 50% 时，取首次日期。'], ['检查异常事件', '比对商品、库存、广告与竞对异常的发生时间。'], ['输出主因', '按链接、购物车、库存、涨价、广告、竞对的优先级给出结论。']], files: [['掉量起点表', '../../04_分析结果/模块B_ASIN掉量起点表.csv'], ['下滑原因归因表', '../../04_分析结果/模块B_下滑原因归因表.csv']] },
  c: { letter: 'C', name: '多 ASIN 活动诊断与重构', summary: '识别适配冲突、搜索词错配与重复投放，形成待运营确认的活动处置清单。', kpiLabel: '活动处置', scope: '广告活动 · 适配冲突与重复投放', steps: [['活动体检', '汇总活动 ASIN 数、适配冲突词、错配花费与订单。'], ['冲突判定', '识别多 ASIN 适配冲突及重复活动。'], ['形成处置', '输出保留、合并、拆出、移除、验证或归档建议。']], files: [['活动处置清单', '../../04_分析结果/模块C_活动处置清单_最终.csv'], ['重复活动清单', '../../04_分析结果/模块C_重复活动清单.csv']] },
  d: { letter: 'D/E', name: '无效流量与否词校验', summary: '模块 D 识别候选否词；模块 E 复核适配范围，避免误伤正确车型、年份或有效流量。', kpiLabel: '候选待确认', scope: '搜索词低效消耗治理', steps: [['识别候选否词', '按不相关品牌、适配外年份等规则生成候选。'], ['适配安全校验', '正确车型或年份命中时必须标记为“禁止否定”。'], ['进入执行队列', '仅“可否定”且经运营确认的词允许实际执行。']], files: [['候选否词清单', '../../04_分析结果/模块D_否词清单.csv'], ['否词校验表', '../../04_分析结果/模块E_否词校验表.csv']] },
  f: { letter: 'F', name: '潜力词', summary: '找出有订单但未有效投放的搜索词，为利润允许的 ASIN 提供小步增投机会。', kpiLabel: '增长机会', scope: '搜索词报告与关键词库', steps: [['识别机会词', '搜索词有订单，但不在当前投放关键词清单中。'], ['排除无效范围', '剔除自家品牌词、已投词与已主动否定词。'], ['推荐增投', '输出推荐 ASIN、匹配方式与 7 天复核动作。']], files: [['潜力词清单', '../../04_分析结果/模块F_潜力词清单.csv']] },
  g: { letter: 'G', name: '盈亏平衡 ACOS', summary: '基于最近完整月损益成本计算平衡 ACOS 与目标 ACOS，是所有增投动作的经营护栏。', kpiLabel: '可增投 ASIN', scope: '有广告动作的 ASIN · 2026-08 成本', steps: [['汇总损益成本', '合并售价、采购、头程、佣金、税费、仓储、Coupon 与广告花费。'], ['计算盈亏线', '计算平衡 ACOS、保守平衡 ACOS 和建议目标 ACOS。'], ['拦截或放行', '高于目标线先治理，处于安全区才允许小步增投。']], files: [['ASIN 盈亏平衡表', '../../04_分析结果/模块G_ASIN盈亏平衡表.csv']] }
};

const esc = value => String(value ?? '—').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
const shorten = value => { const text = String(value ?? '—'); return text.length > 46 ? `${text.slice(0, 46)}…` : text; };
const metricKeys = { a: 'alertAsins', b: 'causeRows', c: 'campaignActions', d: 'negativeCandidates', f: 'opportunities', g: 'investableAsins' };
const requested = new URLSearchParams(location.search).get('module');
const id = Object.hasOwn(modules, requested) ? requested : 'a';
const module = modules[id];
const data = window.WORKBENCH_DATA || {};
const result = data.moduleResults?.[id] || {};
const count = Number(data.metrics?.[metricKeys[id]] ?? 0).toLocaleString();
const logic = result.insights?.length ? result.insights : module.steps.map(step => step[1]);
const distribution = Array.isArray(result.distribution) ? result.distribution : [];
const columns = Array.isArray(result.columns) ? result.columns : [];
const rows = Array.isArray(result.rows) ? result.rows.slice(0, 6) : [];

document.title = `${module.name} · 转向器广告复盘`;
document.querySelectorAll('.nav-link').forEach(link => link.classList.toggle('active', link.getAttribute('href') === `module.html?module=${id}`));
document.querySelector('.workbench-name strong').textContent = '转向器驾驶舱';
document.querySelector('.workbench-name span:last-child').textContent = '陈喆妍 · 周度监控';
document.querySelector('.nav-caption.space').textContent = '监控视图';
const monitorNames = { a: '经营预警', b: '异常归因', c: '投放结构监控', d: '无效流量监控', f: '增长机会监控', g: '利润护栏监控' };
document.querySelectorAll('.left-nav .nav-link[href^="module.html"]').forEach(link => { const key = new URL(link.href).searchParams.get('module'); link.textContent = monitorNames[key] || link.textContent; });

const logicHtml = logic.map((item, index) => `<li><b>${index + 1}</b><span>${esc(item)}</span></li>`).join('');
const distributionHtml = distribution.length
  ? distribution.map(item => `<div class="conclusion-chip ${stateClass(item.label)}"><span>${esc(item.label)}</span><strong>${Number(item.value || 0).toLocaleString()}</strong></div>`).join('')
  : '<p class="empty-note">当前 data.js 尚未同步本模块的结论分布；刷新工作台数据后会自动显示。</p>';
const tableHtml = columns.length && rows.length
  ? `<div class="result-table-wrap"><table class="result-table"><thead><tr>${columns.map(column => `<th>${esc(column[0])}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${columns.map(column => `<td title="${esc(row[column[1]])}">${esc(shorten(row[column[1]]))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`
  : '<div class="empty-note">暂无重点明细。请运行当前模块分析后执行“工作台_刷新数据.py”，页面将显示排序后的前 6 条结果。</div>';
const fileHtml = module.files.map(([label, href]) => `<a class="output" href="${href}"><strong>打开${esc(label)} <span>↗</span></strong><small>本地结果文件 · 以最新分析结果为准</small></a>`).join('');
const flowHtml = [['a', '预警'], ['b', '归因'], ['g', '利润准入'], ['c', '活动重构'], ['d', '否词校验'], ['f', '潜力词']]
  .map(([key, label]) => `<a class="flow-link ${key === id ? 'current' : ''}" href="module.html?module=${key}">${label}</a>`).join('<i>→</i>') + '<i>→</i><a class="flow-link review" href="dashboard.html#review">7 天复核</a>';

document.querySelector('#module-workspace').innerHTML = `
  <section class="module-banner"><div><p class="eyebrow green">广告监控视图 · ${esc(module.scope)}</p><h1>${esc(module.name)}</h1><p>${esc(result.headline || module.summary)}</p></div><div class="module-kpi"><strong>${count}</strong><span>${esc(module.kpiLabel)}</span><small>同步：${esc(data.refreshedAt || '待同步')}</small></div></section>
  <nav class="module-flow" aria-label="广告复盘闭环">${flowHtml}</nav>
  <section class="module-layout"><article class="module-panel"><p class="eyebrow">监控判断逻辑</p><ol class="logic-list">${logicHtml}</ol></article><article class="module-panel"><p class="eyebrow">监控结论分布</p><div class="conclusion-list">${distributionHtml}</div></article></section>
  <section class="module-panel results-panel"><div class="result-heading"><div><p class="eyebrow">重点分析明细</p><h2>前 6 条优先结果</h2><p>用于初步判断；执行前请在完整结果中核对证据与适配范围。</p></div><span class="status-pill">${rows.length ? `已加载 ${rows.length} 条` : '等待数据同步'}</span></div>${tableHtml}</section>
  <section class="module-panel output-panel"><div><p class="eyebrow">完整结果文件</p><h2>在本地打开全部明细</h2></div><div class="output-list">${fileHtml}</div><p class="file-help">使用相对路径以兼容 file:// 本地打开；若链接失效，请确认工作台仍位于资料包的 <code>06_工作台/dist</code> 目录。</p></section>`;

function stateClass(label) {
  const text = String(label || '');
  if (/禁止|异常|直接原因|高于|风险/.test(text)) return 'risk';
  if (/可增投|可否定|增长|正常|安全/.test(text)) return 'safe';
  if (/待|候选|关注|预警/.test(text)) return 'watch';
  return 'neutral';
}
