const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const path = require("path");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  
  const config = {
    mode: dev ? "development" : "production",
    devtool: dev ? "source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/index.tsx", "./src/index.css"]
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash].js",
      clean: true
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".json", ".html"],
      modules: ['node_modules'],
      alias: {
        '@': path.resolve(__dirname, 'src'),
        'components': path.resolve(__dirname, 'src/components'),
        'layout': path.resolve(__dirname, 'src/components/layout')
      },
      fallback: {
        "browser": false
      }
    },
    module: {
      rules: [
        {
          test: /\.(ts|tsx|js|jsx)$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: [
                ["@babel/preset-env", {
                  targets: {
                    ie: "11"
                  },
                  useBuiltIns: "usage",
                  corejs: 3
                }],
                "@babel/preset-typescript",
                ["@babel/preset-react", { "runtime": "automatic" }]
              ],
              plugins: [
                "@babel/plugin-transform-runtime"
              ],
              cacheDirectory: true
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"]
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext]'
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./public/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new webpack.DefinePlugin({
        'process.env': {
          'REACT_APP_API_URL': JSON.stringify(process.env.REACT_APP_API_URL || 'https://backend-rewordLy.onrender.com')
        }
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "../manifest.xml",
            to: "manifest.xml"
          },
          {
            from: "../assets",
            to: "assets",
            noErrorOnMissing: true
          }
        ]
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: '/'
      },
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      port: 3000,
      hot: true
    }
  };

  return config;
}; 