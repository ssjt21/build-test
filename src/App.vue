<script setup lang="ts">
import { ref } from "vue";
import { invoke } from "@tauri-apps/api/core";
import { Folder, Document, Tools, VideoPlay } from "@element-plus/icons-vue";
import { open } from "@tauri-apps/plugin-dialog";
import type { UploadInstance } from "element-plus";
import { ElMessage, ElNotification } from "element-plus";

import PizZip from "pizzip";

// import createReport from "docx-templates";
import Docxtemplater from "docxtemplater";
// import fs from "fs";

// const greetMsg = ref("");
// const name = ref("");
const wordTplPath = ref<string>("");
const saveDir = ref<string>(""); // 文件保存路径
const excelFile = ref<string>("");
// async function greet() {
//   // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
//   greetMsg.value = await invoke("greet", { name: name.value });
// }
// 选择word模版文件
async function chooseWordTpl() {
  // wordTplPath.value = await invoke("choose_word_tpl");
  // Open a dialog
  const file = await open({
    multiple: false,
    directory: false,
    filters: [{ name: "Word", extensions: ["docx"] }],
  });
  console.log(file);
  wordTplPath.value = file;
}

async function chooseSavePath() {
  // wordTplPath.value = await invoke("choose_word_tpl");
  // Open a dialog
  const file = await open({
    multiple: false,
    directory: true,
  });
  console.log(file);
  saveDir.value = file;
}

//自定义处理Excel处理方式
const uploadRef = ref<UploadInstance>();

const submitUpload = () => {
  console.log("uploadRef.value", uploadRef.value);
  // uploadRef.value!.submit();
};

//
const handleChange: UploadProps["onChange"] = (uploadFile, uploadFiles) => {
  // fileList.value = fileList.value.slice(-3);
  // console.log("uploadFile", uploadFile);
  // console.log("uploadFiles", uploadFiles);
  // 获取文件后缀
  const fileExt = uploadFile.name.split(".").pop();
  console.log("fileExt", fileExt);
  if (fileExt !== "xlsx") {
    ElMessage.error("请选择Excel文件[.xlsx]!");
    uploadRef.value!.clearFiles();
    return false;
  }
  //读取文件内容并进行base64编码
  const reader = new FileReader();
  reader.readAsDataURL(uploadFile.raw as File);
  reader.onload = function () {
    const base64Data = reader.result as string;
    excelFile.value = base64Data;
  };
};
const handleRemove: UploadProps["onRemove"] = (uploadFile, uploadFiles) => {
  excelFile.value = null;
  // excelFile.value = uploadFile.raw as File;
  // excelFile.value = uploadFile.raw as File;
};

const beforeFileUpload: UploadProps["beforeUpload"] = (rawFile) => {
  // console.log("rawFile", rawFile);
  //获取文件后缀
  const fileExt = rawFile.name.split(".").pop();
  if (rawFile.type !== "xlsx") {
    ElMessage.error("请选择Excel文件[.xlsx]!");
    return false;
  }
  return true;
};

function toBase64(buffer) {
  var binary = "";
  var bytes = new Uint8Array(buffer);
  var len = bytes.byteLength;

  for (var i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return window.btoa(binary);
}

// submitUpload();
const submitStart = async () => {
  // if (!wordTplPath.value) {
  //   ElMessage.error("请选择Word模版文件!");
  //   return false;
  // }
  // if (!saveDir.value) {
  //   ElMessage.error("请选择文件保存路径!");
  //   return false;
  // }
  // if (!excelFile.value) {
  //   ElMessage.error("请选择Excel文件!");
  //   return false;
  // }
  // console.log(excelFile.value);

  const datalist = await invoke("process_file", {
    docxPath: wordTplPath.value,
    savePath: saveDir.value,
    excelFile: excelFile.value,
  });
  let template = await invoke("get_docx_bs64", { docxPath: wordTplPath.value });
  // console.log("template", template);
  // 将 Base64 字符串转换为二进制缓冲区
  // 将 Base64 字符串转换为二进制数据
  //获取base64字符串
  let bs64Str;
  if (typeof template === "string") {
    bs64Str = template.replace(
      "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,",
      ""
    );
  }
  // interface DataItem {
  //   org_name: string;
  //   org_code: string;
  //   check_month: string;
  //   check_day: string;
  //   rec_month: string;
  //   rec_day: string;
  //   stp_month: string;
  //   stp_day: string;
  //   // Add other properties as needed
  // }
  // console.log("bs64Str", bs64Str);
  const binaryString = atob(bs64Str);
  const len = binaryString.length;
  const binaryData = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    binaryData[i] = binaryString.charCodeAt(i);
  }

  //  // 批量生成文档
  let data_list;
  try {
    data_list = JSON.parse(datalist);
  } catch (error) {
    console.error("Error parsing JSON:", error);
    return;
  }
  // console.log(datalist);

  data_list.forEach(async (item: any, index: number) => {
    // 使用 docx-template 渲染文档
    // const report = await createReport({
    //   template: binaryData.buffer,
    //   data: item,
    // });
    console.log(index);
    const zip = new PizZip(binaryData);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.setData(item);

    try {
      // Render the document
      await doc.render();
    } catch (error) {
      console.error("Error rendering document:", error);
      return;
    }
    const buf = doc.getZip().generate({ type: "arraybuffer" });
    // 将report编码成base64
    let b64ret = toBase64(buf);
    console.log("b64ret", b64ret);
    // 保存结果
    await invoke("save_docx_bs64", {
      docxPath: saveDir.value + "/" + item.org_name + "_整改通知" + ".docx",
      bs64: b64ret,
    });

    // console.log("item", item);
    // result_list.push(b64ret);
  });
  ElNotification({
    title: "处理完成",
    message: "处理完成，请到保存目录查看:" + saveDir.value,
    type: "success",
  });

  // console.log("result_list", result_list);
  // 传回后段处理
  // console.log("result_list", result_list);

  // console.log("ret", JSON.parse(ret));
  // const template = fs.readFileSync(wordTplPath.value, "binary");
  // console.log("template", template);
  // const buffer = await createReport({
  //   template,
  //   data: {
  //     org_name: "test",
  //     surname: "Appleseed",
  //   },
  // });

  // fs.writeFileSync("report.docx", buffer);
};
</script>

<template>
  <!-- <main class="container">
    <h1>文档处理工具（0.1.0）</h1>
  </main> -->
  <div class="common-layout">
    <el-container>
      <el-header>
        <el-text class="mx-1" type="primary" size="large"
          >文档工具v0.1.0</el-text
        ></el-header
      >
      <el-main>
        <div class="mt-4">
          <el-input
            v-model="wordTplPath"
            style="max-width: 500px"
            placeholder="请选择Word模版文件路径"
            class="input-with-select"
          >
            <template #prepend>
              <el-button :icon="Document" />
            </template>
          </el-input>
          <el-button type="primary" @click="chooseWordTpl">
            选择<el-icon class="el-icon--right"><Tools /></el-icon>
          </el-button>
        </div>
        <div class="mt-4">
          <el-input
            v-model="saveDir"
            style="max-width: 500px"
            placeholder="设置文件保存路径"
          >
            <template #prepend>
              <el-button :icon="Folder" />
            </template>
          </el-input>
          <el-button type="primary" @click="chooseSavePath">
            选择<el-icon class="el-icon--right"><Tools /></el-icon>
          </el-button>
        </div>
        <div class="mt-4">
          <el-upload
            class="upload-demo"
            ref="uploadRef"
            drag
            action=""
            :auto-upload="false"
            :limit="1"
            :on-change="handleChange"
            :before-upload="beforeFileUpload"
            :on-remove="handleRemove"
          >
            <el-icon class="el-icon--upload"><upload-filled /></el-icon>
            <div class="el-upload__text">
              请将Exdel文件拖拽到这里<em> 点击选择Excel文件</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">
                <el-text type="danger" size="small"
                  >仅支持（.xlsx）后缀的表格文件</el-text
                >
              </div>
            </template>
          </el-upload>
        </div>
      </el-main>
    </el-container>
  </div>
  <div class="mt-4" style="text-align: center">
    <el-button
      type="primary"
      color="#626aef"
      style="width: 360px"
      @click="submitStart"
    >
      启动<el-icon class="el-icon--right"><VideoPlay /></el-icon>
    </el-button>
  </div>
</template>

<style scoped>
/* .logo.vite:hover {
  filter: drop-shadow(0 0 2em #747bff);
}

.logo.vue:hover {
  filter: drop-shadow(0 0 2em #249b73);
} */
.el-header {
  .el-text {
    font-size: 30px;
    font-weight: 600;
    text-align: center;
  }
}
.el-button {
  margin-left: 10px;
}

.el-container {
  padding: 20px;
}
.el-main {
  align-items: center;
  justify-content: center;
  text-align: center;
  /* height: 100vh; */
}
.mt-4 {
  margin-bottom: 20px;
}
</style>
<style>
/* :root {
  font-family: Inter, Avenir, Helvetica, Arial, sans-serif;
  font-size: 16px;
  line-height: 24px;
  font-weight: 400;

  color: #0f0f0f;
  background-color: #f6f6f6;

  font-synthesis: none;
  text-rendering: optimizeLegibility;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  -webkit-text-size-adjust: 100%;
} */

/* .container {
  margin: 0;
  padding-top: 10vh;
  display: flex;
  flex-direction: column;
  justify-content: center;
  text-align: center;
}

.logo {
  height: 6em;
  padding: 1.5em;
  will-change: filter;
  transition: 0.75s;
}

.logo.tauri:hover {
  filter: drop-shadow(0 0 2em #24c8db);
}

.row {
  display: flex;
  justify-content: center;
}

a {
  font-weight: 500;
  color: #646cff;
  text-decoration: inherit;
}

a:hover {
  color: #535bf2;
}

h1 {
  text-align: center;
}

input,
button {
  border-radius: 8px;
  border: 1px solid transparent;
  padding: 0.6em 1.2em;
  font-size: 1em;
  font-weight: 500;
  font-family: inherit;
  color: #0f0f0f;
  background-color: #ffffff;
  transition: border-color 0.25s;
  box-shadow: 0 2px 2px rgba(0, 0, 0, 0.2);
}

button {
  cursor: pointer;
}

button:hover {
  border-color: #396cd8;
}
button:active {
  border-color: #396cd8;
  background-color: #e8e8e8;
}

input,
button {
  outline: none;
}

#greet-input {
  margin-right: 5px;
}

@media (prefers-color-scheme: dark) {
  :root {
    color: #f6f6f6;
    background-color: #2f2f2f;
  }

  a:hover {
    color: #24c8db;
  }

  input,
  button {
    color: #ffffff;
    background-color: #0f0f0f98;
  }
  button:active {
    background-color: #0f0f0f69;
  }
} */
</style>
