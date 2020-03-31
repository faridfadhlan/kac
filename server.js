const express = require('express');
const app = express();
const port = 3500;
const path = require('path');

app.use('/assets', express.static('assets'))
app.use('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
})

app.listen(port, () => {
    console.log('App started on port ' + port);
})