<template>
  <div class="hello">
    <input type="file" ref="CSVfile" @change="readCSV">
    <button @click="neaten">整理</button>
    <a @click="exportExcel">导出excel</a>
    <table border="1" cellspacing="0">
      <tr>
        <th>ID</th>
        <th>Part Type</th>
        <th>Designator</th>
        <th>Footprint</th>
        <th>Description</th>
        <th>total</th>
      </tr>
      <tr v-for="(item, index) in dataList" :key="index">
        <td>{{index + 1}}</td>
        <td v-for="(value, key) in item" :key="key">{{value}}</td>
      </tr>
      <tr>
        <td colspan="5"></td>
        <td>{{total}}</td>
      </tr>
    </table>
  </div>
</template>

<script>
import XLSX from 'xlsx';
export default {
  name: 'HelloWorld',
  data () {
    return {
      wb: null,
      rABS: false,
      dataList: [],
      total: 0,
      sortRule: []
    }
  },
  methods: {
    readCSV () {
      this.total = 0;
      if (!this.$refs.CSVfile.files[0]) {
        return;
      }
      let f = this.$refs.CSVfile.files[0];
      const reader = new FileReader();
      reader.onload = (e) => {
        let data = e.target.result;
        if (this.rABS) {
          this.wb = XLSX.read(btoa(fixdata(data)), {//手动转化
            type: 'base64'
          });
        } else {
          this.wb = XLSX.read(data, {
            type: 'binary'
          });
          this.dataList = XLSX.utils.sheet_to_json(this.wb.Sheets[this.wb.SheetNames[0]]);
          // this.dataList = this.dataList.concat(XLSX.utils.sheet_to_json(this.wb.Sheets[this.wb.SheetNames[0]]));
          for (var i = 0; i < this.dataList.length; i++) {
            if (!this.dataList[i].Footprint) {
              this.dataList[i].Footprint = '';
            }
            if (!this.dataList[i].Description) {
              this.dataList[i]['Description'] = '';
            }
          }
          console.log(this.dataList);
        }
      };
      if(this.rABS) {
        reader.readAsArrayBuffer(f);
      } else {
        reader.readAsBinaryString(f);
      }
    },
    neaten() {
      // 合并
      for (let i = 0;i < this.dataList.length;) {
        if (i === this.dataList.length -1) {
          if (!this.dataList[i].total) {
            this.dataList[i].total = 1;
          }
          break;
        }
        if (this.dataList[i]['Part Type'] === this.dataList[i + 1]['Part Type'] && this.dataList[i]["Footprint"] === this.dataList[i + 1]["Footprint"] && this.dataList[i]["Description"] === this.dataList[i + 1]["Description"]) {
          this.dataList[i]["Designator"] = this.dataList[i]["Designator"] + ',' + this.dataList[i + 1]["Designator"];
          if (!this.dataList[i].total) {
            this.dataList[i].total = 2;
          } else {
            this.dataList[i].total++;
          }
          // console.log(this.dataList[i].Designator);
          this.dataList.splice(i + 1, 1);
        } else {
          if (!this.dataList[i].total) {
            this.dataList[i].total = 1;
          }
          i++;
        }
      }
      // Designator排序
      this.dataList.forEach((item, index) => {
        this.total += item.total;
        let designatorArr = item.Designator.split(',');
        designatorArr.sort(function(a, b) {
          return (Number(a.replace(/[a-zA-Z]/g, '')) - Number(b.replace(/[a-zA-Z]/g, '')));
        });
        let result = designatorArr.join(',');
        console.log(result);
        this.dataList[index].Designator = result;
      });
      //  Description排序
    },
    exportExcel() {
      // 使用outerHTML属性获取整个table元素的HTML代码（包括<table>标签），然后包装成一个完整的HTML文档，设置charset为urf-8以防止中文乱码
        var html = "<html><head><meta charset='utf-8' /></head><body>" + document.getElementsByTagName("table")[0].outerHTML + "</body></html>";
        // 实例化一个Blob对象，其构造函数的第一个参数是包含文件内容的数组，第二个参数是包含文件类型属性的对象
        var blob = new Blob([html], { type: "application/vnd.ms-excel" });
        var a = document.getElementsByTagName("a")[0];
        // 利用URL.createObjectURL()方法为a元素生成blob URL
        a.href = URL.createObjectURL(blob);
        // 设置文件名，目前只有Chrome和FireFox支持此属性
        a.download = "整理结果.xls";
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
</style>
