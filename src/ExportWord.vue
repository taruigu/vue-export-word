<template>
  <div id="app">
    <vue-line :data="chartData"
              :extend="chartExtend"
              :after-set-option="afterSetOption"></vue-line>
    <div>
      <button @click.self="exportCharts">导出</button>
    </div>
  </div>
</template>

<script>
import { saveAs } from 'file-saver'
import PizZip from 'pizzip'
import PizZipUtils from "pizzip/utils"
import docxtemplater from "docxtemplater"
import ImageModule from "docxtemplater-image-module-free"
// import echarts from "echarts"
// import VCharts from "v-charts"
import VeLine from 'v-charts/lib/line.common'

export default {
  data: function () {
    return {
      chartData: {
        columns: ['月份', '工作量'],
        rows: [
          { '月份': '4月', '工作量': 2 },
          { '月份': '5月', '工作量': 5 },
          { '月份': '6月', '工作量': 20 },
          { '月份': '7月', '工作量': 17 },
          { '月份': '8月', '工作量': 25 },
          { '月份': '9月', '工作量': 22 }
        ]
      },
      reportData: {
        "author": "小卷毛塔瑞古",
        "content": "2020-05-10-report-0096"
      },
      base64DataURL: "",
      chartExtend: {
        // 设置关闭图表生成的动态效果，否则导出的图片只有坐标图，没有柱和线
        animation: false
      }
    }
  },
  components: {
    "vue-line": VeLine
  },
  methods: {
    /**
     * v-charts没有直接提供echarts的getDataURL()方法
     * 这里用钩子函数触发原生的getDataURL()方法，并将返回值绑定到定义好的变量上使用
     */
    afterSetOption: function (chartObj) {
      this.base64DataURL = chartObj.getDataURL();
    },
    /**
     * 此方法为docxtemplater-image-module-free官网提供
     */
    base64DataURLToArrayBuffer: function (dataURL) {
      const base64Regex = /^data:image\/(png|jpg|svg|svg\+xml);base64,/;
      if (!base64Regex.test(dataURL)) {
        return false;
      }
      const stringBase64 = dataURL.replace(base64Regex, "");
      let binaryString;
      if (typeof window !== "undefined") {
        binaryString = window.atob(stringBase64);
      } else {
        binaryString = new Buffer(stringBase64, "base64").toString("binary");
      }
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        const ascii = binaryString.charCodeAt(i);
        bytes[i] = ascii;
      }
      return bytes.buffer;
    },
    exportCharts: function () {
      const ___this = this;
      // 读取模板，如果要读取静态图片，方式一样
      // 模板必须放在public文件夹下，如果有图片，也一样
      PizZipUtils.getBinaryContent("word4export.docx", function (error, content) {
        if (error) {
          throw error;
        }
        // 图表导出，本质还是图片的导出
        let opts = {}
        opts.centered = true;  // 图片居中
        opts.getImage = function (chartURL) {
          return ___this.base64DataURLToArrayBuffer(chartURL);
        }
        /**
         *  三个参数，第一个是图片对象，第二个在这里是base64DataURL，第三个是名称, 如果多个图表可根据第三个判断，
         *  给定相应的大小，或者修改docxtemplater-image-module-free官网的，方法，动态的给定大小值
         */

        opts.getSize = function () {
          return [600, 300]; //大小可以动态调整，或者根据不同的图，给定不同的大小
        }
        let imageModule = new ImageModule(opts);
        // 创建一个PizZip实例，内容为模板的内容
        let zip = new PizZip(content);
        // 创建并加载docxtemplater实例对象
        let doc = new docxtemplater();
        doc.attachModule(imageModule);
        doc.loadZip(zip);

        // 设置模板变量的值
        doc.setData({
          "chart-pic": ___this.base64DataURL, // v-chart图表数据
          ...___this.reportData
        });

        try {
          doc.render(); // 渲染文档
        } catch (error) {
          console.log(error);
        }
        // 生成导出文件，这些都是docxtemplater-image-module-free官网的方法
        let out = doc.getZip().generate({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        saveAs(out, "exportWord.docx"); // 导出文件，给定文件名
      });
    }
  }
}
</script>

<style scoped>
vue-line {
  width: 600px;
  height: 300px;
}
</style>