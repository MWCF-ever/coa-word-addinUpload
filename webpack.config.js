const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const fs = require('fs');

module.exports = (env, argv) => {
  const isDev = argv.mode === 'development';
  const isLocalDev = isDev && process.env.NODE_ENV !== 'production';
  
  // HTTPS证书配置 - 只在本地开发时使用
  let httpsOptions = false;
  if (isLocalDev && fs.existsSync(path.join(__dirname, './certs/localhost.key'))) {
    httpsOptions = {
      key: fs.readFileSync(path.join(__dirname, './certs/localhost.key')),
      cert: fs.readFileSync(path.join(__dirname, './certs/localhost.crt'))
    };
  }

  return {
    entry: {
      taskpane: './src/taskpane/index.tsx',
      commands: './src/commands/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: isDev ? '[name].js' : '[name].[contenthash].js',
      clean: true,
      // 修复：使用相对路径而不是绝对路径
      publicPath: './'
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx']
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            {
              loader: 'babel-loader',
              options: {
                presets: [
                  '@babel/preset-env',
                  '@babel/preset-react',
                  '@babel/preset-typescript'
                ]
              }
            }
          ],
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.(png|svg|jpg|jpeg|gif)$/i,
          type: 'asset/resource'
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/index.html',
        filename: 'taskpane.html',
        chunks: ['taskpane'],
        // 确保使用相对路径
        inject: 'body'
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'manifest.xml',
            to: 'manifest.xml'
          },
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true
          }
        ]
      })
    ],
    // devServer配置只在本地开发时有效
    devServer: isLocalDev ? {
      hot: true,
      https: httpsOptions,
      port: 3000,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
        'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
      },
      static: {
        directory: path.join(__dirname, 'dist'),
        // 重要：使用相对路径
        publicPath: '/'
      }
    } : undefined,
    optimization: {
      splitChunks: {
        chunks: 'all',
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            priority: 10
          }
        }
      }
    },
    // 性能警告配置 - 提高限制或禁用警告
    performance: {
      hints: isDev ? false : 'warning',
      maxAssetSize: 2000000, // 2MB
      maxEntrypointSize: 2000000 // 2MB
    }
  };
};