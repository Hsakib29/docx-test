export default `
.dt{--ink:#16222e;--mute:#5d6b79;--line:#d5dce3;--bg:#f6f8fa;--card:#fff;--acc:#0b6b78;--acc-ink:#fff;--bad:#a8261d;
  font-family:"Hind Siliguri","Noto Sans Bengali",system-ui,sans-serif;color:var(--ink);max-width:960px;margin:0 auto;padding:24px 16px;line-height:1.5}
@media (prefers-color-scheme:dark){.dt{--ink:#e7edf2;--mute:#9aa8b5;--line:#2e3a46;--bg:#10171d;--card:#17212a;--acc:#4cc0cf;--acc-ink:#06242a;--bad:#ff8a80}}
.dt *{box-sizing:border-box}
.dt h2{margin:0 0 4px;font-size:1.6rem}
.dt-lead{margin:0 0 20px;color:var(--mute);max-width:62ch}
.dt-bar{display:flex;flex-wrap:wrap;gap:10px 14px;align-items:center;margin:14px 0}
.dt label{display:flex;gap:6px;align-items:center;font-size:.9rem}
.dt select,.dt button{font:inherit;font-size:.9rem;padding:7px 11px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--ink)}
.dt button{cursor:pointer}
.dt button:hover:not(:disabled){border-color:var(--acc)}
.dt button:disabled{opacity:.5;cursor:not-allowed}
.dt .dt-primary{background:var(--acc);color:var(--acc-ink);border-color:var(--acc);font-weight:600}
.dt :focus-visible{outline:2px solid var(--acc);outline-offset:2px}
.dt-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px}
@media (max-width:640px){.dt-grid{grid-template-columns:1fr}}
.dt-drop{display:flex;flex-direction:column;gap:4px;padding:22px 16px;border:2px dashed var(--line);border-radius:10px;background:var(--card);cursor:pointer;text-align:center}
.dt-drop span{color:var(--mute);font-size:.88rem}
.dt-drop em{font-style:normal;font-size:.85rem;word-break:break-word}
.dt-drop.is-over,.dt-drop:hover{border-color:var(--acc);background:var(--bg)}
.dt-mem{margin-left:auto;color:var(--mute);font-size:.9rem}
.dt-progress{margin:10px 0;font-size:.88rem;color:var(--mute)}
.dt-track{height:6px;background:var(--line);border-radius:3px;overflow:hidden;margin-bottom:4px}
.dt-track div{height:100%;background:var(--acc);transition:width .2s}
.dt-note{font-size:.88rem;color:var(--mute);margin:8px 0}
.dt-result{border:1px solid var(--line);border-radius:10px;background:var(--card);margin-top:14px;overflow:hidden}
.dt-head{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center;padding:12px 16px}
.dt-head div:first-child{display:flex;flex-direction:column;min-width:0}
.dt-head span{font-size:.85rem;color:var(--mute)}
.dt-head .dt-err{color:var(--bad)}
.dt-actions{display:flex;gap:8px}
.dt-scroll{overflow-x:auto;border-top:1px solid var(--line)}
.dt-scroll table{width:100%;border-collapse:collapse;font-size:.88rem}
.dt-scroll th,.dt-scroll td{text-align:left;vertical-align:top;padding:8px 12px;border-bottom:1px solid var(--line);white-space:pre-wrap}
.dt-scroll th{background:var(--bg);font-weight:600}
.dt-tag{font-size:.78rem;padding:2px 8px;border-radius:99px;border:1px solid var(--line);white-space:nowrap}
.dt-tag.dt-tm{border-color:var(--acc);color:var(--acc)}
.dt-tag.dt-error{border-color:var(--bad);color:var(--bad)}
.dt-tag.dt-ok{border-color:var(--acc);color:var(--acc);margin-left:4px}
.dt-scroll tr[data-pending="true"] td:first-child{box-shadow:inset 3px 0 0 var(--acc)}
.dt-tag.dt-edited{background:var(--acc);border-color:var(--acc);color:var(--acc-ink)}
.dt-seg{display:inline-flex}
.dt .dt-seg button{border-radius:0;margin-left:-1px}
.dt .dt-seg button:first-child{border-radius:6px 0 0 6px;margin-left:0}
.dt .dt-seg button:last-child{border-radius:0 6px 6px 0}
.dt .dt-seg .is-on{background:var(--ink);color:var(--card);border-color:var(--ink);position:relative}
.dt-filter{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:8px 16px;border-top:1px solid var(--line)}
.dt-filter span{margin-left:auto;font-size:.82rem;color:var(--mute)}
.dt .dt-filter .is-on{background:var(--ink);color:var(--card);border-color:var(--ink)}
.dt-scroll textarea{width:100%;min-width:240px;font:inherit;color:inherit;background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:6px 8px;resize:vertical}
.dt-scroll textarea:focus{background:var(--card);border-color:var(--acc)}
.dt .dt-mini{display:block;margin-top:6px;padding:3px 8px;font-size:.78rem}
.dt-top{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
.dt-top h2{margin:0}
.dt-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-top:14px}
.dt-card{border:1px solid var(--line);border-radius:10px;background:var(--card);padding:14px 16px;display:flex;flex-direction:column;gap:8px}
.dt-card h3{margin:0;font-size:1.05rem}
.dt-meta{color:var(--mute);font-size:.85rem}
.dt-empty{margin-top:18px;padding:22px;border:2px dashed var(--line);border-radius:10px;color:var(--mute)}
.dt-form{display:grid;gap:14px;padding:16px;border:1px solid var(--line);border-radius:10px;background:var(--card);margin-top:14px;max-width:680px}
.dt-form label{display:flex;flex-direction:column;align-items:flex-start;gap:4px;font-size:.9rem}
.dt-form fieldset{border:1px solid var(--line);border-radius:8px;padding:8px 12px}
.dt-form fieldset label{flex-direction:row;align-items:flex-start;gap:8px;padding:4px 0}
.dt-form fieldset small{display:block;color:var(--mute)}
.dt .dt-form input[type=text],.dt-gl input[type=text],.dt-gl input[type=search]{font:inherit;font-size:.9rem;padding:7px 10px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink);width:100%}
.dt-tabs{display:flex;flex-wrap:wrap;gap:4px;border-bottom:1px solid var(--line);margin:16px 0}
.dt .dt-tabs button{border:0;border-bottom:2px solid transparent;border-radius:0;background:none;padding:8px 14px}
.dt .dt-tabs .is-on{border-bottom-color:var(--acc);font-weight:600}
.dt-terms{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
.dt-term{font-size:.78rem;padding:2px 8px;border-radius:99px;border:1px solid var(--line)}
.dt-term.ok{border-color:var(--acc);color:var(--acc)}
.dt-term.bad{border-color:var(--bad);color:var(--bad)}
.dt-gl-add{display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:8px;margin:12px 0}
@media (max-width:640px){.dt-gl-add{grid-template-columns:1fr}}
.dt-gl .dt-scroll{border:1px solid var(--line);border-radius:8px}
.dt-gl td{padding:4px 6px}
`;
