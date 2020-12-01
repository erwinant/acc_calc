<template>
  <q-page padding>
    <q-file v-model="files" label="Pick files"
    accept=".xlsx,.xlsm,xls,csv" outlined use-chips
    @input="onFileChange">
      <template v-slot:prepend>
        <q-icon name="attach_file" />
      </template>
    </q-file>
  </q-page>
</template>

<script>
import XLSX from 'xlsx'
export default {
  data() {
    return {
      files: []
    };
  },
  methods: {
    onFileChange(file){
      if(file){
      let promise = new Promise((resolve) => {
        let reader = new FileReader()
        reader.onload = e => {
          let data = e.target.result;
          let workbook = XLSX.read(data, {
            type: "binary",
          });
          workbook.SheetNames.forEach(function (sheetName) {
            // Here is your object
            let XL_row_object = XLSX.utils.sheet_to_row_object_array(
              workbook.Sheets[sheetName]
            );
            let json_object = JSON.stringify(XL_row_object);
            console.log(json_object);
          });
        }
        reader.onerror = function (ex) {
          console.log(ex);
        };
        reader.readAsBinaryString(file);
      })
      }
    },
    reader() {
      let reader = new FileReader();
      reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: "binary",
        });

        workbook.SheetNames.forEach(function (sheetName) {
          // Here is your object
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(
            workbook.Sheets[sheetName]
          );
          var json_object = JSON.stringify(XL_row_object);
          console.log(json_object);
        });
      };

      reader.onerror = function (ex) {
        console.log(ex);
      };

      reader.readAsBinaryString(file);
    },
  },
};
</script>
