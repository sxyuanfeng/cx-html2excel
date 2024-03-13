/*
 * @Author: xujiang
 * @Date: 2024-03-13 15:31:40
 * @LastEditors: xujiang
 * Copyright (c) 2024 by xujiang/cicc, All Rights Reserved.
 */
const path = require('path')

function buildConfig(prod, umd = false) {
  const config = {
    mode: 'development',
    entry: {
      'cx-html2excel': path.join(__dirname, './src/html2excel.js')
    },
    output: {
      path: path.join(__dirname, './dist'),
      filename: `[name]${prod ? '.min' : ''}.${umd ? '' : 'm'}js`,
      globalObject: 'globalThis'
    },
    devtool: 'source-map',
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: [{ loader: 'ts-loader' }]
        }
      ]
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    externals: {
      "exceljs": {
        root: "ExcelJS",
        commonjs: "exceljs",
        commonjs2: "exceljs",
        amd: "exceljs",
        module: 'exceljs'
      },
    }
  };

  if (umd) {
    config.output.library = { name: 'html2excel', type: 'umd', umdNamedDefine: true };
  } else {
    config.experiments = { outputModule: true };
    config.output.library = { type: 'module' };
  }

  return config;
}


module.exports = (env, argv) => {
  return buildConfig(argv.mode === 'production', env.umd);
};
