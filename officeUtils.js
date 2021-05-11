import Axios from "axios";

// 由于医院的用户浏览器版本太低，使用前端导出表格不是乱码就会有各种各样的兼容性问题，因此写了一个服务器小工具专门用来生成表格，
// 请用blob来接受响应，如果出错会响应text/plain错误提示，如果没问题会响应application/vnd.ms-excel文件流
//请求的格式为（post）
// {
//   colWidth: 40,
//   rowHeight: 25,
//   head: ["课程体系包名称", "课程体系类型", "套餐类型", "已分配学习卡"],
//   data: [
//    ["护理服务", "活动赠送课程", "1选1", 55],
//    ["天津海河医院-核心能力建设", "NCC课程", "35选6", 5],
//    ["Test Excel", "NCC课程", "100选6", 5]
//   ],
//   headStyle: JSON.stringify({
//     alignment: { horizontal: "center", vertical: "center" },
//     font: { bold: true },
//     fill: { type: "pattern", color: ["#cbcbcb"], pattern: 1 }
//    }),
//   dataStyle: JSON.stringify({
//    alignment: { vertical: "center" }
//    })
// }
const instance = Axios.create({
  baseURL: "http://localhost:8086",
  timeout: 60000,
  headers: {
    Accept: "*/*"
  },
  responseType: "blob"
});

const exportExcel = (data, fileName = "") => {
  instance
    .post("/build_excel", data)
    .then(r => {
      let blob = r.data;
      if (blob.type === "text/plain") {
        blob.text().then(r => {
          alert(r);
        });
        return;
      } else if (blob.type === "application/vnd.ms-excel") {
        const aLink = document.createElement("a");
        aLink.href = URL.createObjectURL(blob);
        aLink.setAttribute("download", fileName + ".xlsx"); // 设置下载文件名称
        document.body.appendChild(aLink);
        aLink.click();
        return;
      }
      alert("未知数据");
    })
    .catch(err => {
      alert(JSON.stringify(err));
    });
};

const buildData = (head = [], list = [], config) => {
  let data = {
    colWidth: 15,
    rowHeight: 25,
    headStyle: JSON.stringify({
      alignment: { horizontal: "center", vertical: "center" },
      font: { bold: true },
      fill: { type: "pattern", color: ["#cbcbcb"], pattern: 1 }
    }),
    dataStyle: JSON.stringify({
      alignment: { vertical: "center" }
    })
  };
  Object.assign(data, config);
  data.head = [];
  data.data = [];

  head.forEach(h => {
    data.head.push(h.title);
  });
  list.forEach(r => {
    let row = [];
    head.forEach(f => {
      if (f.hasOwnProperty("func")) {
        row.push(f.func(r));
      } else {
        row.push(r[f.field]);
      }
    });
    data.data.push(row);
  });
  return data;
};

const downloadExcel = (head = [], list = [], fileName = "", config) => {
  exportExcel(buildData(head, list, config), fileName);
};

export { downloadExcel };
