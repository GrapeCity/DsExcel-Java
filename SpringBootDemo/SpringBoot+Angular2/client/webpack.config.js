const path = require('path');
const CheckerPlugin = require('awesome-typescript-loader').CheckerPlugin;
const bundleOutputDir = '../src/main/resources/public';
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        'polyfills': './src/polyfills.ts',
        'bundle': './src/main.ts'
    },
    output: {
        filename: "[name].js",
        path: path.join(__dirname, bundleOutputDir)
    },

    // Enable sourcemaps for debugging webpack's output.
    devtool: "source-map",

    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: [".ts", ".tsx", ".js", ".json"]
    },
    optimization: {
        minimize: true
    },
    module: {
        rules: [
            { test: /\.ts$/, loaders: ['awesome-typescript-loader', 'angular2-template-loader'] },
            { test: /\.html$/, use: 'html-loader?minimize=false' },
            { test: /\.css$/, use: [ 'to-string-loader', 'css-loader'] },
            { test: /\.(png|jpg|jpeg|gif|svg)$/, use: 'url-loader?limit=25000' }
        ]
    },
    plugins: [
        new CheckerPlugin(),
        new CopyWebpackPlugin([
            { from: './index.html' },
            { from: './favicon-16x16.png' },
            { from: './favicon-32x32.png' },
            { from: './src/spread.sheets', to: 'spreadJS' },
            { from: './src/css', to: 'css' }
        ], {
                ignore: [
                    '*.scss'
                ],
                copyUnmodified: true
            })
    ]
};