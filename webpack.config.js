var config = {
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
};

var pluginConfig = Object.assign({}, config, {
    entry: './client/src/index.js',
    output: {
        filename: './client/fin.desktop.Excel.js'
    }
});

var loaderConfig = Object.assign({}, config, {
    entry: './provider/src/provider.js',
    output: {
        filename: './provider/provider.js'
    }
});

module.exports = [
    pluginConfig,
    loaderConfig
];