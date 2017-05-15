module.exports = {
    entry: './src/main.js',
    output: {
        filename: './web/excel-api-example.js'
    },
    resolve: {
        extensions: ['.js']
    },
    module: {
        loaders: [
            {
                exclude: /(node_modules|bower_components)/
            }
        ]
    }
}