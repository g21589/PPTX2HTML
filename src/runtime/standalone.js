/** Entry point embedded in generated HTML documents. */
import { autoStart, Player } from './player.js';

window.PPTXPlayer = Player;
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', autoStart);
else autoStart();
