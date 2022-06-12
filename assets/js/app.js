var uploadedFile = document.getElementById("excelFile");
let selectedFile;
uploadedFile.addEventListener("change", (event) => {
  selectedFile = event.target.files[0];
});
var downloadBtn = document.getElementById("download-json");
downloadBtn.style.display = "none";


// Convert Function Excel2JSON
var convertBtn = document.getElementById("buttonConvert");
convertBtn.addEventListener("click", (event) => {
  if (selectedFile) {
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(selectedFile);
    fileReader.onload = (event) => {
      let data = event.target.result;
      let workbook = XLSX.read(data, { type: "binary" });
      var product = {};
      var d = [];
      workbook.SheetNames.forEach((sheet, index) => {
        let rowObject = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
        rowObject.forEach((prod, i) => {
          d.push({
            id: prod["id"],
            catalog_number: prod["catalog_number"],
            new_arrivals_flg: prod["new_arrivals_flg"],
            limited_edition_flg: prod["limited_edition_flg"],
            drop_flg: prod["drop_flg"],
            product_name_ja: prod["product_name_ja"],
            product_name_en: prod["product_name_en"],
            category: prod["category"],
            price_no_tax: prod["price_no_tax"],
            item_description: prod["item_description"],
            image: prod["image"],
            detail_link: prod["detail_link"],
          });
        });
        product["productList"] = d;
        var convertedJson = document.getElementById("displayJson");
        convertedJson.innerHTML = JSON.stringify(
          product,
          null,
          4
        );
        if (convertedJson.value != null) {
          document.getElementById("download-json").style.display = "block";
        }
      });
    };
  }
});

// Download Converted JSON to Local Machine
downloadBtn.addEventListener("click", (event) => {
  var jsonStr = document.getElementById("displayJson").value;
  downloadObjectAsJson(jsonStr, "productList");
  function downloadObjectAsJson(str, filename) {
    var d_str =
      "data:text/json;charset=utf-8," +
      encodeURIComponent(str) +
      ";base64," +
      encodeURIComponent(str);
    var anchor = document.createElement("a");
    anchor.setAttribute("href", d_str);
    anchor.setAttribute("download", filename + ".json");
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }
});
