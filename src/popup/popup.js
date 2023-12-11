/* global Office self window global console*/
console.info("_sample【popup】boot");

Office.onReady(() => {});

function closeDialog(val) {
  console.info("_sample【popup】sendMessage", val)
  window.close()
}

window.onbeforeunload = () => {
  window.opener[window.name](this.sendFlag)
  self.close()
},
window.addEventListener('message', (event) => {
  if (event.origin === window.location.origin) {
    console.log(event.data)
  }
},false)

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.closeDialog = closeDialog;

