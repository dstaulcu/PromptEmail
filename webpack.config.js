const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');
const packageJson = require('./package.json');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.js',
    commands: './src/commands/commands.js'
  },
  output: {
    filename: '[name].bundle.js',
    path: path.resolve(__dirname, 'public'),
    clean: true
  },
  module: {
    rules: [
      {
        test: /\.css$/i,
        use: ['style-loader', 'css-loader']
      },
      {
        test: /\.(png|svg|jpg|jpeg|gif)$/i,
        type: 'asset/resource',
        generator: {
          filename: 'icons/[name][ext]'
        }
      }
    ]
  },
  plugins: [
    new webpack.DefinePlugin({
      'process.env.PACKAGE_VERSION': JSON.stringify(packageJson.version)
    }),
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      template: './src/commands/commands.html',
      filename: 'index.html',
      chunks: ['commands']
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './src/assets/icons',
          to: 'icons'
        },
        {
          from: './src/assets/css',
          to: '.',
          globOptions: {
            ignore: ['**/*.scss']
          }
        },
        {
          from: './src/config/ai-providers.json',
          to: 'config/ai-providers.json'
        },
        {
          from: './src/config/prompts.json',
          to: 'config/prompts.json'
        },
        {
          from: './src/config/taskpane-resources.json',
          to: 'config/taskpane-resources.json'
        },
        {
          from: './src/config/telemetry.json',
          to: 'config/telemetry.json'
        }
      ]
    })
  ],
  resolve: {
    extensions: ['.js', '.css']
  },
  devtool: 'source-map',
  optimization: {
    minimize: true
  }
};
