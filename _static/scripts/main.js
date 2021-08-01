"use strict";function _typeof(t){return(_typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t})(t)}!function(t,e){"function"==typeof define&&define.amd?define([],function(){return e(t)}):"object"===("undefined"==typeof exports?"undefined":_typeof(exports))?module.exports=e(t):t.Gumshoe=e(t)}("undefined"!=typeof global?global:"undefined"!=typeof window?window:void 0,function(l){function u(t,e,n){n.settings.events&&(n=new CustomEvent(t,{bubbles:!0,cancelable:!0,detail:n}),e.dispatchEvent(n))}function n(t){var e=0;if(t.offsetParent)for(;t;)e+=t.offsetTop,t=t.offsetParent;return 0<=e?e:0}function d(t){t&&t.sort(function(t,e){return n(t.content)<n(e.content)?-1:1})}function i(t,e,n){return t=t.getBoundingClientRect(),e="function"==typeof(e=e).offset?parseFloat(e.offset()):parseFloat(e.offset),n?parseInt(t.bottom,10)<(l.innerHeight||document.documentElement.clientHeight):parseInt(t.top,10)<=e}function r(){return Math.ceil(l.innerHeight+l.pageYOffset)>=Math.max(document.body.scrollHeight,document.documentElement.scrollHeight,document.body.offsetHeight,document.documentElement.offsetHeight,document.body.clientHeight,document.documentElement.clientHeight)}function m(t,e){var n,o,s=t[t.length-1];if(n=s,o=e,!(!r()||!i(n.content,o,!0)))return s;for(var c=t.length-1;0<=c;c--)if(i(t[c].content,e))return t[c]}function o(t,e){e.nested&&t.parentNode&&((t=t.parentNode.closest("li"))&&(t.classList.remove(e.nestedClass),o(t,e)))}function p(t,e){var n;!t||(n=t.nav.closest("li"))&&(n.classList.remove(e.navClass),t.content.classList.remove(e.contentClass),o(n,e),u("gumshoeDeactivate",n,{link:t.nav,content:t.content,settings:e}))}function v(t,e){!e.nested||(t=t.parentNode.closest("li"))&&(t.classList.add(e.nestedClass),v(t,e))}var y={navClass:"active",contentClass:"active",nested:!1,nestedClass:"active",offset:0,reflow:!1,events:!0};return function(t,e){var n,s,c,o,i,r={setup:function(){n=document.querySelectorAll(t),s=[],Array.prototype.forEach.call(n,function(t){var e=document.getElementById(decodeURIComponent(t.hash.substr(1)));e&&s.push({nav:t,content:e})}),d(s)}};r.detect=function(){var t,e,n,o=m(s,i);o?c&&o.content===c.content||(p(c,i),e=i,!(t=o)||(n=t.nav.closest("li"))&&(n.classList.add(e.navClass),t.content.classList.add(e.contentClass),v(n,e),u("gumshoeActivate",n,{link:t.nav,content:t.content,settings:e})),c=o):c&&(p(c,i),c=null)};function a(){o&&l.cancelAnimationFrame(o),o=l.requestAnimationFrame(r.detect)}function f(){o&&l.cancelAnimationFrame(o),o=l.requestAnimationFrame(function(){d(s),r.detect()})}r.destroy=function(){c&&p(c,i),l.removeEventListener("scroll",a,!1),i.reflow&&l.removeEventListener("resize",f,!1),i=o=c=n=s=null};return i=function(){var n={};return Array.prototype.forEach.call(arguments,function(t){for(var e in t){if(!t.hasOwnProperty(e))return;n[e]=t[e]}}),n}(y,e||{}),r.setup(),r.detect(),l.addEventListener("scroll",a,!1),i.reflow&&l.addEventListener("resize",f,!1),r}});
"use strict";var tocScroll=null,header=null;function scrollHandlerForHeader(){0==Math.floor(header.getBoundingClientRect().top)?header.classList.add("scrolled"):header.classList.remove("scrolled")}function scrollHandlerForTOC(e){null!==tocScroll&&(0==e?tocScroll.scrollTo(0,0):Math.ceil(e)>=Math.floor(document.documentElement.scrollHeight-window.innerHeight)?tocScroll.scrollTo(0,tocScroll.scrollHeight):document.querySelector(".scroll-current"))}function scrollHandler(e){scrollHandlerForHeader(),scrollHandlerForTOC(e)}function setTheme(e){"light"!==e&&"dark"!==e&&"auto"!==e&&(console.error("Got invalid theme mode: ".concat(e,". Resetting to auto.")),e="auto"),document.body.dataset.theme=e,localStorage.setItem("theme",e),console.log("Changed to ".concat(e," mode."))}function cycleThemeOnce(){var e=localStorage.getItem("theme")||"auto";setTheme(window.matchMedia("(prefers-color-scheme: dark)").matches?"auto"===e?"light":"light"==e?"dark":"auto":"auto"===e?"dark":"dark"==e?"light":"auto")}function setupScrollHandler(){var o,t=!1;window.addEventListener("scroll",function(e){o=window.scrollY,t||(window.requestAnimationFrame(function(){scrollHandler(o),t=!1}),t=!0)}),window.scroll()}function setupScrollSpy(){null!==tocScroll&&new Gumshoe(".toc-tree a",{reflow:!0,recursive:!0,navClass:"scroll-current"})}function setupTheme(){setTheme(localStorage.getItem("theme")||"auto");var e=document.getElementsByClassName("theme-toggle");Array.from(e).forEach(function(e){e.addEventListener("click",cycleThemeOnce)})}function setup(){setupTheme(),setupScrollHandler(),setupScrollSpy()}function main(){document.body.parentNode.classList.remove("no-js"),header=document.querySelector("header"),tocScroll=document.querySelector(".toc-scroll"),setup()}document.addEventListener("DOMContentLoaded",main);
//# sourceMappingURL=main.js.map
