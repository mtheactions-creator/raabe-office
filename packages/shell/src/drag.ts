/**
 * Drags that follow the mouse through window listeners (mousemove, mouseup) end with the button's
 * release. When another window takes the mouse during a drag (it comes to the front), that release
 * never reaches the page and the drag would go on following a mouse whose button is up. A move
 * without a button after a press therefore stands in for the missing release: a "mouseup" is sent
 * at that place first, so every such drag ends there (D-127).
 */
export function installLostMouseUp(): void {
  let down = false;
  addEventListener("mousedown", (e) => { if (e.button === 0) down = true; }, true);
  addEventListener("mouseup", (e) => { if (e.button === 0) down = false; }, true);
  addEventListener("mousemove", (e) => {
    if (!down || e.buttons & 1) return;
    down = false;
    dispatchEvent(new MouseEvent("mouseup", { clientX: e.clientX, clientY: e.clientY, screenX: e.screenX, screenY: e.screenY, button: 0, buttons: 0, bubbles: true }));
  }, true);
}
