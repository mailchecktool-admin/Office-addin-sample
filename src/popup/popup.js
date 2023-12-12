/* global Office self window global console*/
console.info("_sample【popup】boot");

Office.onReady((info) => {console.log(info)})
//Initialization not performed, There is no console output either

const g = getGlobal();

function closeDialog(val) {
  console.info("_sample【popup】sendMessage", val)
  g.close()
}

g.onbeforeunload = () => {
  g.opener[window.name](true)
  g.close()
},
g.addEventListener('message', (event) => {
  if (event.origin === g.location.origin) {
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

// The add-in command functions need to be available in global scope
g.closeDialog = closeDialog;

