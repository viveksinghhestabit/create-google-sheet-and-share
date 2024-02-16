const { WebSocketServer } = require('ws');

const socket = (server) => {
    const wss = new WebSocketServer({ server });
    const clients = [];

    wss.on('connection', (ws) => {
        clients.push(ws);
        ws.on('message', (message) => {
            console.log(message);
            clients.forEach((client) => {
                if (client.readyState === WebSocket.OPEN) {
                    client.send(message);
                }
            });
        });
    });
};

module.exports = socket;