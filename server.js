var express = require('express')
var path = require('path')
var upload = require('./routes/upload')

var app = express()

app.set('views', path.join(__dirname, 'views'))
app.set('view engine', 'ejs')

app.use('/',upload)
app.use('/static',express.static(path.join(__dirname, 'public')))


const port = process.env.port || 8088; //对应一个端口号

app.listen(port,() =>{

    console.log(`Server running on port ${port}`);

})