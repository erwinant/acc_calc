<template>
  <div>
    <q-toolbar class="bg-secondary text-white">
      <q-toolbar-title>STO ASSET ACC</q-toolbar-title>
    </q-toolbar>
    <div class="q-ma-sm q-pa-xl">
      <q-list bordered round>
        <q-expansion-item
          header-class="bg-grey-3"
          group="somegroup"
          default-opened
          icon="filter_1"
          label="Merge OAT"
          caption="Pick multiple OAT"
        >
          <q-separator />
          <q-card>
            <q-card-section>
              <div class="row q-col-gutter-md q-ma-lg">
                <div class="col-12">
                  <q-file
                    v-model="files"
                    label="Pick files"
                    multiple
                    accept=".xlsx, .xlsm, xls, csv"
                    outlined
                    use-chips
                  >
                    <template v-slot:prepend>
                      <q-icon name="attach_file" />
                    </template>
                  </q-file>
                </div>
                <div class="col-12 row justify-center q-col-gutter-lg q-mt-md">
                  <div class="col-3"></div>
                  <div class="col-6">
                    <q-btn
                      icon="publish"
                      class="full-width"
                      size="lg"
                      unelevated
                      color="primary"
                      label="load file"
                      @click="onLoadStep1"
                    />
                  </div>
                  <div class="col-3"></div>
                  <div class="col-3"></div>
                  <div class="col-6">
                    <download-excel
                      v-if="this.dataExportStep1 !=null"
                      :data="this.dataExportStep1.json_data"
                      :fields="this.dataExportStep1.json_fields"
                      name="merge_step_1.xlsx"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download"
                        unelevated
                        color="teal"
                        icon="get_app"
                      />
                    </download-excel>
                  </div>
                  <div class="col-3"></div>
                </div>
              </div>
            </q-card-section>
          </q-card>
        </q-expansion-item>
        <q-expansion-item
          header-class="bg-grey-3"
          group="somegroup"
          icon="filter_2"
          label="Lookup AR vs OAT"
          caption="Upload AR & OAT">
          <q-separator />
          <q-card>
            <q-card>
            <q-card-section>
              <div class="row q-col-gutter-md  q-ma-lg">
                <div class="col-12">
                  <q-file
                    v-model="fileAR"
                    label="Pick files AR"
                    accept=".xlsx, .xlsm, .xls, .csv"
                    outlined
                    use-chips
                  >
                    <template v-slot:prepend>
                      <q-icon name="attach_file" />
                    </template>
                  </q-file>
                </div>
                <div class="col-12">
                  <q-file
                    v-model="fileOAT"
                    label="Pick files OAT Merged"
                    accept=".xlsx, .xlsm, .xls, .csv"
                    outlined
                    use-chips
                  >
                    <template v-slot:prepend>
                      <q-icon name="attach_file" />
                    </template>
                  </q-file>
                </div>
                <div class="col-12 row justify-center q-col-gutter-lg q-mt-md">
                  <div class="col-3"></div>
                  <div class="col-6">
                    <q-btn
                      icon="publish"
                      class="full-width"
                      size="lg"
                      unelevated
                      color="primary"
                      label="load file"
                      @click="onLoadStep2"
                    />
                  </div>
                  <div class="col-3"></div>
                  <div class="col-3"></div>
                  <div class="col-6">
                    <q-btn
                      icon="play_arrow"
                      class="full-width"
                      size="lg"
                      unelevated text-color="black"
                      color="yellow"
                      label="Compare!"
                      @click="compareStep2"
                      v-if="readyCompare"
                    />
                  </div>
                  <div class="col-3"></div>
                  <div class="col-3"></div>
                  <div class="col-6">
                    <download-excel
                      v-if="this.resultARStep2.length >0"
                      :data="this.resultARStep2"
                      :fields="this.resultARStep2_fields"
                      name="merge_step_1.xlsx"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download AR"
                        unelevated
                        color="teal"
                        icon="get_app"
                      />
                    </download-excel>
                  </div>
                  <div class="col-3"></div>
                  <div class="col-3"></div>
                  <div class="col-6">
                    <download-excel
                      v-if="this.resultOATStep2.length >0"
                      :data="this.resultOATStep2"
                      :fields="this.resultOATStep2_fields"
                      name="merge_step_1.xlsx"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download OAT"
                        unelevated
                        color="teal"
                        icon="get_app"
                      />
                    </download-excel>
                  </div>
                  <div class="col-3"></div>
                </div>
              </div>
            </q-card-section>
          </q-card>
          </q-card>
        </q-expansion-item>
      </q-list>
    </div>
  </div>
</template>
<script>
import { mapActions, mapMutations } from "vuex";
import XLSX from "xlsx";
import { date } from "quasar";
export default {
  data() {
    return {
      files: [],
      dataStep1: [],
      headerStep1: [],
      dataExportStep1: null,
      fileAR:null,
      fileOAT:null,
      dataAR:[],
      dataOAT:[],
      readyCompare:false,
      resultARStep2:[],
      resultOATStep2:[],
      resultARStep2_fields:{},
      resultOATStep2_fields:{}
    };
  },
  methods: {
    async onLoadStep1() {
      if (this.files) {
        let counter = 0;
        for (let i = 0; i < this.files.length; i++) {
          let result = await this.onFileChangeStep1(this.files[i]);
          let data = result.data.map((m) => {
            return {
              ba: result.baCode,
              asset_no: m["F-ST"],
              room_sap: m["__EMPTY"],
              description: m["__EMPTY_1"],
              room_opname: m["__EMPTY_2"],
              asset_status_code: m["__EMPTY_3"],
              tanggal_scan: date.formatDate(m["__EMPTY_4"], "YYYY-MM-DD"),
              waktu_scan: date.formatDate(m["__EMPTY_5"], "HH:mm"),
              keterangan: m["__EMPTY_6"] || "",
              input_manual: m["__EMPTY_7"],
            };
          });

          this.dataStep1 = [...this.dataStep1, ...data];
          this.headerStep1 = result.header;
          if (i == this.files.length - 1) {
            let json_fields = {};
            for (const [key, value] of Object.entries(this.dataStep1[0])) {
              json_fields[key] = key;
            }
            this.dataExportStep1 = {
              json_fields: json_fields,
              json_data: this.dataStep1,
              json_meta: [
                [
                  {
                    key: "charset",
                    value: "utf-8",
                  },
                ],
              ],
            };
            console.log(this.dataStep1);
          }
        }
      }
    },
    async onFileChangeStep1(file) {
      if (file) {
        return new Promise((resolve) => {
          let reader = new FileReader();
          reader.onload = (e) => {
            let data = e.target.result;
            let workbook = XLSX.read(data, {
              type: "binary",
              sheets: 1,
              cellDates: true,
            });

            workbook.SheetNames.forEach(function (sheetName) {
              let XL_row_object = XLSX.utils.sheet_to_row_object_array(
                workbook.Sheets[sheetName]
              );
              if (sheetName === "FORM OAT") {
                let baCode = XL_row_object[1]["__EMPTY_1"]
                  ? XL_row_object[1]["__EMPTY_1"]
                  : "";
                let header = XL_row_object[4] ? XL_row_object[4] : "";
                let fileData = XL_row_object.slice(5, XL_row_object.length);
                // let json_object = JSON.stringify(XL_row_object.slice(4,XL_row_object.length));
                resolve({
                  baCode: baCode,
                  data: fileData,
                  header: header,
                });
              }
            });
          };
          reader.onerror = function (ex) {
            console.log(ex);
          };
          reader.readAsBinaryString(file);
        });
      }
    },
    async onLoadStep2() {
      if(this.fileAR && this.fileOAT){
          this.$q.loading.show({
            message: 'Please wait'
          })
          let objectAR = await this.onFileChangeStep2(this.fileAR, 'AR')
          let objectOAT = await this.onFileChangeStep2(this.fileOAT, 'OAT')
          if(objectAR.length == 0 || objectOAT.length == 0){
            this.$q.dialog({message:"AR or OAT incorrect content", title:"Incomplete File"})
            return
          }else{
            this.dataAR = [...objectAR]
            this.dataOAT = [...objectOAT]
            this.readyCompare = true
            this.$q.loading.hide()
          }
      }else{
        this.$q.dialog({message:"Please upload AR and OAT file", title:"Incomplete File"})
      }
    },
    parser(type, data){
      if(type ==='AR'){
        return data.map(m=>{
                return {
                  asset :m.Asset,
                  asset_description :m['Asset description'],
                  bus_a :m.BusA,
                  asset :m.Asset,
                  cap_date :date.formatDate(m['Cap.date'], "YYYY-MM-DD"),
                  class :m.Class,
                  or_asset :m['Or. asset'],
                  s_no :m['SNo.'],
                  book_val :m['      Book val.'],
                  accum_dep :m['     Accum.dep.'],
                  acquis_val :m['    Acquis.val.'],
                  deact_date :date.formatDate(m['Deact.Date'], "YYYY-MM-DD") || '',
                  concate : `${m.Asset} ${m['SNo.']}`
                }
              })
      }
      if(type ==='OAT'){
        return data.map(m=>{
          return {
            ...m,
            tanggal_scan:date.formatDate(m.tanggal_scan, "YYYY-MM-DD") || '',
          }
        })
      }
    },
    compareStep2(){
      this.$q.loading.show({
        message: 'Please wait'
      })
      setTimeout(() => {
        this.resultARStep2 = this.dataAR.map(m=>{
          let tanggal_sto = this.dataOAT.find(f=>f.asset_no === m.concate) ? '' : this.dataOAT[5].tanggal_scan
          return {
            ...m,
            find_on_oat: this.dataOAT.find(f=>f.asset_no === m.concate) ? '' : 'N/A',
            tanggal_sto : tanggal_sto,
            check : m.cap_date > tanggal_sto ? 'BELI' : 'CEK'
          }
        })
        for (const [key, value] of Object.entries(this.resultARStep2[5])) {
            this.resultARStep2_fields[key] = key;
        }
        this.resultOATStep2 = this.dataOAT.map(m=>{
          return {
            ...m,
            find_on_ar: this.dataAR.find(f=>f.concate === m.asset_no) ? '' : 'N/A',
            find_mutasi:  this.dataAR.find(f=>f.or_asset === m.asset_no) ?  this.dataAR.find(f=>f.or_asset === m.asset_no).bus_a : 'N/A'
          }
        })
        for (const [key, value] of Object.entries(this.resultOATStep2[5])) {
            this.resultOATStep2_fields[key] = key;
        }
        this.$q.loading.hide()
      }, 500);

    },
    async onFileChangeStep2(file, type) {
      if (file) {
        return new Promise((resolve) => {
          let reader = new FileReader();
          reader.onload = (e) => {
            let data = e.target.result;
            let workbook = XLSX.read(data, {
              type: "binary",
              sheets: 0,
              cellDates: true,
            });

            workbook.SheetNames.forEach((sheetName) =>{
              let XL_row_object = XLSX.utils.sheet_to_row_object_array(
                workbook.Sheets[sheetName]
              );
              resolve(this.parser(type, XL_row_object))
            });
          };
          reader.onerror = function (ex) {
            console.log(ex);
          };
          reader.readAsBinaryString(file);
        });
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
    tranformCsv(input, filename) {
      let csvContent = "data:text/csv;charset=utf-8,";
      csvContent += [
        Object.keys(input[0]).join(";"),
        ...input.map((item) => Object.values(item).join(";")),
      ].join("\n");

      const data = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", data);
      link.setAttribute("download", `${filename}.csv`);
      link.click();
    },
  },
};
</script>