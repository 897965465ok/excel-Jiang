<template>
  <div id="excel">
    <input type="file" name="file" class="file" @change="getFile"/>
  </div>
</template>
<script >
import "@/xspreadsheet.js";
import "@/zh-cn.js";
import "@/xspreadsheet.css";
import "@/hint.min.css";
import excel from "./excel.js";
export default {
  name: "Sheet",
  async mounted() {
    this.initData();
  },
  data() {
    return {
      sheet: null,
      hostyData: null
    };
  },
  // activated、deactivated
  methods: {
    async initData() {
      x_spreadsheet.locale("zh-cn");
      this.sheet = x_spreadsheet("#excel");
      let localFile = document.querySelector(".top-0 > div");
      let fileButton = document.querySelector(".file");
      let exportLocalFile = document.querySelector(".export-file");
      let user = document.querySelector(".user");
      let userLiset = document.querySelectorAll(".user div");
      let top4 = document.querySelector(".top-4");
      this.tips(top4, "文档类型");
      this.tips(localFile, "导入本地文件");
      this.tips(exportLocalFile, "导出到本地");
      user.onclick = () => {
        if (user.isShow) {
          userLiset[0].style.display = "block";
          userLiset[1].style.display = "none";
          user.isShow = false;
          this.sheet.loadData(this.hostyData).reRender();
        } else {
          userLiset[1].style.display = "block";
          userLiset[0].style.display = "none";
          user.isShow = true;
          this.hostyData = this.sheet.getData();
          this.sheet.loadData({}).reRender();
        }
      };

      localFile.onclick = () => {
        fileButton.click();
      };
      exportLocalFile.onclick = () => {
        this.exportFile();
      };
    },
    async loadingFile() {
      let opting = await excel.loadExel();
      this.sheet.loadData([opting]).reRender();
    },
    async loadBase64() {
      let img = {};
      if (img && img.table) {
        let workbook = await XlsxPopulate.fromDataAsync(
          Buffer.from(img.table, "base64")
        );
        let opting = await excel.loadExel(workbook);
        this.sheet.loadData(opting).reRender();
      }
    },
    async exportFile() {
      await excel.exportFile(this.sheet.getData());
    },
    async getFile(evnent) {
      let opting = await excel.loadExel(null, evnent.target.files[0]);
      this.sheet.loadData([opting]).reRender();
    },
    tips(element, message) {
      element.classList.add("hint--bottom");
      element.setAttribute("aria-label", message);
    }
  }
};
</script>
<style lang="scss" scoped>
.file {
  display: none;
}
</style>