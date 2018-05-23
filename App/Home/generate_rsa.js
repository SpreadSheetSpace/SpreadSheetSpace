importScripts('../../Scripts/cryptico.js');

function generateRsa(username) {
    var rsa = cryptico.generateRSAKey(username, 2048);
    postMessage(rsa);
}

self.addEventListener('message', function (event) {
    var username = event.data;
    generateRsa(username);
}, false);