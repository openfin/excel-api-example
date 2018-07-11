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
    entry: './src/plugin.js',
    output: {
        filename: './plugin/fin.desktop.Excel.js'
    }
});

var loaderConfig = Object.assign({}, config, {
    entry: './src/service-loader.js',
    output: {
        filename: './plugin/service-loader.js'
    }
});

module.exports = [
    pluginConfig,
    loaderConfig
];