(() => {
  if (window.top !== window) return;
  const pages = [
    ['劳务用工分析', 'labor-funnel-prototype-v2.html', '劳务用工需求履约分析'],
    ['劳务用工分析', 'dashboard.html', '劳务用工分析驾驶舱'],
    ['劳务用工分析', 'journey.html', '用户旅程与模型'],
    ['劳务用工分析', '用工分析报表PRD.html', '用工分析报表 PRD'],
    ['组件优化', 'department-tree-select.html', '部门树下拉选单'],
    ['工作代理 · OTWS', '工作代理.html', '工作代理需求清单'],
    ['工作代理 · OTWS', 'hrone-work-agent-prototype/index.html', '工作代理用户旅程'],
    ['工作代理 · OTWS', 'hrone-work-agent-prototype/portal.html', '系统管理 · 工作代理'],
    ['工作代理 · OTWS', 'hrone-work-agent-prototype/logs.html', '工作代理日志'],
    ['试工管理', '门户-试工管理.html', '门户-试工评估'],
    ['试工管理', 'otws-试工管理.html', 'OTWS-试工评估'],
    ['试工管理', '门户-派遣需求.html', '门户-派遣需求'],
    ['试工管理', '门户-派遣人员名单.html', '门户-派遣人员名单'],
    ['试工管理', '需求处理规则.html', 'OTWS-需求处理规则'],
    ['消息提醒', 'message-reminder-prototype/config.html?v=nav-message-reminder-63', '消息提醒配置'],
    ['消息提醒', 'message-reminder-prototype/journey.html?v=nav-message-reminder-61', '消息提醒用户旅程'],
    ['消息提醒', 'message-reminder-prototype/消息提醒-PRD.html?v=nav-message-reminder-62', '消息提醒 PRD'],
    ['消息提醒', 'mentoring-rule-prototype/index.html?revision=6#rules', '带教规则配置'],
    ['环节工时登记和报表初步方案', '扫码工时登记方案/海外仓劳务工扫码工时登记完整方案.html', '海外仓环节工时登记与报表方案']
  ];
  const scriptUrl = document.currentScript?.src;
  if (!scriptUrl || document.getElementById('prototype-switcher')) return;
  const rootUrl = new URL('.', scriptUrl);
  const style = document.createElement('style');
  style.textContent = `
    #prototype-switcher{position:fixed;z-index:2147483000;width:54px;height:54px;touch-action:none;user-select:none}
    #prototype-switcher-trigger{display:grid;place-items:center;width:54px;height:54px;padding:0;overflow:hidden;border:2px solid #fff;border-radius:50%;background:#fff;box-shadow:0 3px 14px #18212f40;cursor:grab}
    #prototype-switcher-trigger:active{cursor:grabbing}
    #prototype-switcher-trigger:focus-visible{outline:3px solid #1888e8;outline-offset:2px}
    #prototype-switcher-trigger img{width:100%;height:100%;object-fit:cover;pointer-events:none}
    #prototype-switcher-panel{position:fixed;z-index:2147483000;width:min(300px,calc(100vw - 24px));max-height:min(420px,70vh);display:flex;flex-direction:column;overflow:hidden;border:1px solid #d7dce3;border-radius:8px;background:#fff;box-shadow:0 12px 36px #10182838;color:#202833;font:13px/1.45 -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif}
    #prototype-switcher-panel[hidden]{display:none}
    .ps-head{display:flex;align-items:center;justify-content:space-between;min-height:44px;padding:0 12px;border-bottom:1px solid #e8ebef;font-weight:650}
    .ps-close{width:30px;height:30px;border:0;border-radius:4px;background:transparent;color:#566171;font-size:20px;cursor:pointer}
    .ps-search{margin:10px 12px 8px;width:calc(100% - 24px);height:34px;padding:0 9px;border:1px solid #cfd6df;border-radius:5px;font:inherit;outline:none}
    .ps-search:focus{border-color:#2689e8;box-shadow:0 0 0 2px #2689e826}
    .ps-list{overflow:auto;padding:0 6px 8px}
    .ps-group{margin:0 3px;border-bottom:1px solid #edf0f3}
    .ps-group>summary{display:flex;align-items:center;gap:7px;min-height:32px;padding:5px 7px;color:#344054;font-weight:650;cursor:pointer;list-style:none}
    .ps-group>summary::-webkit-details-marker{display:none}
    .ps-group>summary:before{content:"";width:0;height:0;border:4px solid transparent;border-left-color:#7a8594;transform:translateY(1px)}
    .ps-group:not([open])>summary:before{transform:translateY(1px)}
    .ps-group[open]>summary:before{transform:rotate(90deg) translateX(1px)}
    .ps-group{border-bottom:1px solid #edf0f3}
    .ps-link{margin-left:14px}
    .ps-link{display:block;padding:8px 9px;border-radius:5px;color:#344054;text-decoration:none;overflow-wrap:anywhere}
    .ps-link:hover,.ps-link:focus-visible{background:#f1f6fb;color:#126db2;outline:none}
    .ps-link[aria-current=page]{background:#eaf4fd;color:#126db2;font-weight:600}
    .ps-empty{padding:14px 9px;color:#707b89}
    @media(max-width:480px){#prototype-switcher{width:48px;height:48px}#prototype-switcher-trigger{width:48px;height:48px}#prototype-switcher-panel{max-height:65vh}}
  `;
  document.head.append(style);

  const widget = document.createElement('div');
  widget.id = 'prototype-switcher';
  const trigger = document.createElement('button');
  trigger.id = 'prototype-switcher-trigger';
  trigger.type = 'button';
  trigger.setAttribute('aria-label', '切换原型页面');
  trigger.setAttribute('aria-haspopup', 'dialog');
  trigger.setAttribute('aria-expanded', 'false');
  const avatar = document.createElement('img');
  avatar.src = new URL('prototype-avatar.png', rootUrl).href;
  avatar.alt = '';
  trigger.append(avatar);
  widget.append(trigger);

  const panel = document.createElement('section');
  panel.id = 'prototype-switcher-panel';
  panel.hidden = true;
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-label', '切换原型页面');
  const head = document.createElement('div');
  head.className = 'ps-head';
  const heading = document.createElement('span');
  heading.textContent = '原型页面';
  const close = document.createElement('button');
  close.className = 'ps-close';
  close.type = 'button';
  close.setAttribute('aria-label', '关闭');
  close.textContent = '×';
  head.append(heading, close);
  const search = document.createElement('input');
  search.className = 'ps-search';
  search.type = 'search';
  search.placeholder = '搜索页面';
  search.setAttribute('aria-label', '搜索原型页面');
  const list = document.createElement('nav');
  list.className = 'ps-list';
  list.setAttribute('aria-label', '原型页面列表');
  panel.append(head, search, list);
  document.body.append(widget, panel);

  const current = new URL(location.href);
  const groupNodes = new Map();
  const links = pages.map(([groupName, href, title]) => {
    let group = groupNodes.get(groupName);
    if (!group) {
      group = document.createElement('details');
      group.className = 'ps-group';
      group.open = true;
      const summary = document.createElement('summary');
      summary.textContent = groupName;
      group.append(summary);
      groupNodes.set(groupName, group);
      list.append(group);
    }
    const link = document.createElement('a');
    link.className = 'ps-link';
    link.href = new URL(href, rootUrl).href;
    link.textContent = title;
    link.dataset.search = `${title} ${href} ${groupName}`.toLocaleLowerCase();
    link.dataset.group = groupName;
    if (new URL(link.href).pathname === current.pathname) {
      link.setAttribute('aria-current', 'page');
      link.addEventListener('click', event => event.preventDefault());
    }
    group.append(link);
    return link;
  });

  const positionKey = 'otws.prototype-switcher.position.v1';
  let position = { x: innerWidth - 74, y: innerHeight - 78 };
  try {
    const saved = JSON.parse(localStorage.getItem(positionKey));
    if (Number.isFinite(saved?.x) && Number.isFinite(saved?.y)) position = saved;
  } catch {}
  const place = () => {
    const size = trigger.offsetWidth || 54;
    position.x = Math.max(8, Math.min(position.x, innerWidth - size - 8));
    position.y = Math.max(8, Math.min(position.y, innerHeight - size - 8));
    widget.style.left = `${position.x}px`;
    widget.style.top = `${position.y}px`;
    if (!panel.hidden) placePanel();
  };
  const placePanel = () => {
    const width = panel.offsetWidth || 300;
    const height = panel.offsetHeight || 300;
    panel.style.left = `${Math.max(12, Math.min(position.x + 54 - width, innerWidth - width - 12))}px`;
    panel.style.top = `${Math.max(12, Math.min(position.y + 54 - height, innerHeight - height - 12))}px`;
  };
  place();
  addEventListener('resize', place);

  const closePanel = () => {
    panel.hidden = true;
    trigger.setAttribute('aria-expanded', 'false');
    trigger.focus();
  };
  const openPanel = () => {
    panel.hidden = false;
    trigger.setAttribute('aria-expanded', 'true');
    placePanel();
    search.value = '';
    filterLinks();
    search.focus();
  };
  trigger.addEventListener('click', () => {
    if (dragged) return;
    panel.hidden ? openPanel() : closePanel();
  });
  close.addEventListener('click', closePanel);
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && !panel.hidden) closePanel();
  });
  document.addEventListener('pointerdown', event => {
    if (!panel.hidden && !panel.contains(event.target) && !widget.contains(event.target)) {
      panel.hidden = true;
      trigger.setAttribute('aria-expanded', 'false');
    }
  });
  function filterLinks() {
    const query = search.value.trim().toLocaleLowerCase();
    let count = 0;
    links.forEach(link => {
      link.hidden = !link.dataset.search.includes(query);
      if (!link.hidden) count++;
    });
    groupNodes.forEach(group => {
      const hasVisible = [...group.querySelectorAll('.ps-link')].some(link => !link.hidden);
      group.hidden = !hasVisible;
      if (query && hasVisible) group.open = true;
    });
    let empty = list.querySelector('.ps-empty');
    if (!count && !empty) {
      empty = document.createElement('div');
      empty.className = 'ps-empty';
      empty.textContent = '没有匹配的页面';
      list.append(empty);
    } else if (count && empty) empty.remove();
  }
  search.addEventListener('input', filterLinks);

  let drag = null;
  let dragged = false;
  trigger.addEventListener('pointerdown', event => {
    if (event.button !== 0) return;
    drag = { id: event.pointerId, x: event.clientX, y: event.clientY, ox: position.x, oy: position.y };
    dragged = false;
    trigger.setPointerCapture(event.pointerId);
  });
  trigger.addEventListener('pointermove', event => {
    if (!drag || event.pointerId !== drag.id) return;
    const dx = event.clientX - drag.x;
    const dy = event.clientY - drag.y;
    if (Math.abs(dx) + Math.abs(dy) > 4) dragged = true;
    if (dragged) {
      position.x = drag.ox + dx;
      position.y = drag.oy + dy;
      place();
    }
  });
  trigger.addEventListener('pointerup', () => {
    if (!drag) return;
    drag = null;
    try { localStorage.setItem(positionKey, JSON.stringify(position)); } catch {}
    if (dragged) setTimeout(() => { dragged = false; }, 0);
  });
  trigger.addEventListener('pointercancel', () => { drag = null; dragged = false; });
})();
