const path = require('path');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const CheckerPlugin = require('awesome-typescript-loader').CheckerPlugin;
const bundleOutputDir = '../src/main/resources/public';
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: "./src/boot.tsx",
    output: {
        filename: "bundle.js",
        path: path.join(__dirname, bundleOutputDir)
    },

    // Enable sourcemaps for debugging webpack's output.
    devtool: "source-map",

    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: [".ts", ".tsx", ".js", ".json"]
    },
    optimization:{
        minimize: true
    },
    module: {
        rules: [
            // All files with a '.ts' or '.tsx' extension will be handled by 'awesome-typescript-loader'.
            { test: /\.tsx?$/, loader: "awesome-typescript-loader" },

            // All output '.js' files will have any sourcemaps re-processed by 'source-map-loader'.
            { enforce: "pre", test: /\.js$/, loader: "source-map-loader" },

            { test: /\.css$/, use: ExtractTextPlugin.extract({ use: 'css-loader?minimize' }) }
        ]
    },
    plugins: [
        new CheckerPlugin(),
        new CopyWebpackPlugin([
            { from: './index.html' },
            { from: './favicon-16x16.png' },
            { from: './favicon-32x32.png' },
            { from: './src/spread.sheets', to: 'spreadJS' },
            { from: './src/css/vendor.css', to: 'css' }
          ], {
              ignore: [
                '*.scss'
              ],
              copyUnmodified: true
            })
    ].concat([
        new ExtractTextPlugin('css/site.css')
    ])
};