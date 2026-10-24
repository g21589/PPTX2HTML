/** Base stylesheet for rendered slides. */
export const BASE_CSS = `
.pp-deck{position:relative}
.pp-slide{position:relative;overflow:hidden;background:#fff;color:#000;text-align:left;line-height:normal;font-size:16px;contain:layout paint;-webkit-text-size-adjust:none;text-size-adjust:none}
.pp-slide *{box-sizing:border-box}
.pp-bg{position:absolute;left:0;top:0;right:0;bottom:0}
.pp-bg svg{display:block}
.pp-el{position:absolute;transform-origin:50% 50%;margin:0;padding:0}
.pp-geom{position:absolute;left:0;top:0;overflow:visible;pointer-events:none;display:block}
.pp-tx{position:absolute;display:flex;flex-direction:column;overflow:visible;word-wrap:break-word;overflow-wrap:break-word;transform-origin:50% 50%}
.pp-p{margin:0;padding:0;font-size:0;white-space:pre-wrap;flex:none;text-indent:0}
.pp-p span{white-space:inherit}
.pp-lsb{line-height:calc(var(--pp-l) * var(--pp-fs,1))}
.pp-lsb span,.pp-lsb .pp-bu{vertical-align:bottom;line-height:min(1.2em,calc(var(--pp-l) * var(--pp-fs,1)))}
.pp-nowrap .pp-p{white-space:pre;width:max-content;min-width:100%}
.pp-actr{align-items:center}
.pp-actr-in{max-width:100%}
.pp-bu{display:inline-block;text-indent:0;white-space:pre;font-style:normal;font-weight:normal;text-decoration:none}
.pp-bu-gap{font-size:0.25em;white-space:pre}
.pp-p .pp-bu-w{white-space:nowrap}
.pp-a{color:inherit;text-decoration:none}
.pp-hit{position:absolute;left:0;top:0;width:100%;height:100%;z-index:1;display:block;background:transparent}
.pp-table{border-collapse:collapse;table-layout:fixed;border-spacing:0}
.pp-table td{overflow:visible;word-wrap:break-word;overflow-wrap:break-word}
.pp-cell-diag{position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none}
.pp-media video{position:absolute;left:0;top:0;width:100%;height:100%;object-fit:fill;display:block;background:#000}
.pp-media .pp-media-icon{position:absolute;left:0;top:0;width:100%;height:100%}
.pp-media-btn{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:44px;height:44px;border-radius:22px;border:0;background:rgba(0,0,0,.55);color:#fff;font-size:18px;cursor:pointer;opacity:0;transition:opacity .2s}
.pp-media:hover .pp-media-btn,.pp-media.pp-paused .pp-media-btn{opacity:1}
.pp-media.pp-playing .pp-media-btn{opacity:0}
.pp-chart svg{display:block;overflow:visible}
.pp-slide [data-slide],.pp-slide [data-jump],.pp-slide [data-href]{cursor:pointer}
`;
