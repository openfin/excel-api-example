module.exports = {
    entry: './src/plugin.js',
    output: {
        filename: './plugin/fin.desktop.Excel.js'
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