/* global Office self window global console */
console.info("_sample【index】boot");

Office.onReady(() => {})

function onSend(event) {
  console.info("_sample【index】send ckick")
  const wSize = 1250
  const hSize = 650
  let callbackId = 'ID'
  let dialog = window.open('https://localhost:3000/popup.html',callbackId,'width=' + wSize + ',height=' + hSize)
  if (dialog) {
    window.addEventListener('message', () => {
      dialog.onload = () => {
        dialog.postMessage('postMessage', window.location.origin)
      }
    })
  } else {
    event.completed({ allowEvent: false })
  }
  window[callbackId] = async (val) => {
    if(dialog) {
      event.completed({ allowEvent: val })
    } else if(!dialog || dialog.closed || typeof dialog.closed == "undefined") {
      event.completed({ allowEvent: false })
    }
  }
}

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
g.onSend = onSend;
