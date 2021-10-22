const devCerts = require("office-addin-dev-certs");
const CopyPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require('webpack');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react'
      ],
      taskpane: [
        'react-hot-loader/patch',
        './src/taskpane/taskpane.ts',
      ],
      onsend: './src/onsend/onsend.ts',
      dialog: './src/onsend/dialog.tsx'
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            'react-hot-loader/webpack',
            'ts-loader'
          ],
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        filename: "onsend.html",
        template: "./src/onsend/onsend.html",
        chunks: ["onsend"]
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/onsend/dialog.html",
        chunks: ["dialog"]
      }),
      new CopyPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: './assets',
            to: 'assets',
          },
          {
            from: './manifest.xml',
            to: 'manifest.xml',
          }
        ]
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"]
      })
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
