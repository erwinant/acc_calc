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
                  <div class="col-3">
                    <download-excel
                      v-if="this.dataExportStep1 !=null"
                      :data="this.dataExportStep1.json_data"
                      :fields="this.dataExportStep1.json_fields"
                      name="merge_step_1.xls"
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
                  <div class="col-3">
                    <download-excel
                      v-if="this.dataExportStep1 !=null"
                      :data="this.dataExportStep1.json_data.filter(f=>f.keterangan !=='')"
                      :fields="this.dataExportStep1.json_fields"
                      name="merge_step_1_keterangan_filtered.xls"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download Keterangan"
                        outline
                        color="teal"
                        icon="filter_alt"
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
          label="Lookup AR01 vs OAT"
          caption="Upload AR01 & OAT">
          <q-separator />
          <q-card>
            <q-card>
            <q-card-section>
              <div class="row q-col-gutter-md  q-ma-lg">
                <div class="col-12">
                  <q-file
                    v-model="fileAR"
                    label="Pick files AR01"
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
                      name="merge_step_1_ar.xls"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download AR01"
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
                      name="merge_step_2_oat.xls"
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
        <q-expansion-item
          header-class="bg-grey-3"
          group="somegroup"
          icon="filter_3"
          label="Grouping By Asset Status"
          caption="Upload OAT">
          <q-separator />
          <q-card>
            <q-card>
            <q-card-section>
              <div class="row q-col-gutter-md  q-ma-lg">
                <div class="col-12">
                  <q-file
                    v-model="fileOATMerged"
                    label="Pick files OAT"
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
                      @click="onLoadStep3"
                    />
                  </div>
                  <div class="col-3"></div>
                  <div class="col-3"></div>
                  <div class="col-6">
                    <download-excel
                      v-if="this.resultPivotStep3.length >0"
                      :data="this.resultPivotStep3"
                      :fields="this.resultPivotStep3_fields"
                      name="merge_step_3_pivoted.xls"
                    >
                      <q-btn
                        class="full-width"
                        size="lg"
                        label="Download Grouped"
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
      fileOATMerged:null,
      dataAR:[],
      dataOAT:[],
      dataPivoted:[],
      readyCompare:false,
      readyPivoted:false,
      resultARStep2:[],
      resultOATStep2:[],
      resultARStep2_fields:{},
      resultOATStep2_fields:{},

      resultPivotStep3:[],
      resultPivotStep3_fields:{},
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
              waktu_scan: m["__EMPTY_5"],
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
                workbook.Sheets[sheetName], {raw:false}
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
    async onLoadStep3() {
      const groupBy = (value = [], field = "") => {
          const groupedObj = value.reduce((prev, cur) => {
              if (!prev[cur[field]]) {
                  prev[cur[field]] = [cur];
              } else {
                  prev[cur[field]].push(cur);
              }
              return prev;
          }, {});
          return Object.keys(groupedObj).map(key => ({ key, value: groupedObj[key] }));
      }
      if(this.fileOATMerged){
          this.$q.loading.show({
            message: 'Please wait'
          })
          let objectOAT = await this.onFileChangeStep2(this.fileOATMerged, 'OAT')
          if(objectOAT.length == 0){
            this.$q.dialog({message:"OAT incorrect content", title:"Incomplete File"})
            return
          }else{
            this.dataPivoted = [...objectOAT]
            let ba_distinct = [...new Set(this.dataPivoted.map(m=>m.ba))]
            this.readyPivoted = true
            this.$q.loading.hide()
            this.dataPivoted = this.dataPivoted.map(m=>{
              return{
                ...m,
                group_1:`${m.ba}|${m.input_manual?'INPUT_MANUAL':'NONE'}`,
                group_2:`${m.ba}|${m.asset_status_code}`
              }
            })
            let gp1 = groupBy(this.dataPivoted,'group_1').map(m=>{
              return {
                ...m,
                ba:m.key.split('|')[0],
                fields:m.key.split('|')[1]
              }
            }).filter(f=>f.fields==='INPUT_MANUAL').map(m=>{
              return {
                ba:m.ba,
                fields:m.fields,
                value:m.value
              }
            })
            let gp2 = groupBy(this.dataPivoted,'group_2').map(m=>{
              return {
                ...m,
                ba:m.key.split('|')[0],
                fields:m.key.split('|')[1]
              }
            }).map(m=>{
              return {
                ba:m.ba,
                fields:m.fields,
                value:m.value
              }
            })
            ba_distinct = ba_distinct.map(m=>{
              let newObj = { ba_code:m }
              let temp = [...gp1,...gp2].filter(f=>f.ba === m).map(i=>{
                let col_name = i.fields.split('-').join('_').split(' ').join('_')
                newObj[col_name] =i.value.length
                return {[col_name] : i.value.length}
              })
              return { ...newObj}
            })
            let get_prop = [ba_distinct.map(el=>{
              return {
                ...el,
                count_prop:Object.keys(el).length
              }
            }).sort((b,a) => (a.count_prop > b.count_prop) ? 1 : ((b.count_prop > a.count_prop) ? -1 : 0))[0]]
            for (const [key, value] of Object.entries(get_prop[0])) {
              if(key!=="count_prop")
                this.resultPivotStep3_fields[key] = key

            }
            this.resultPivotStep3 = ba_distinct
          }
      }else{
        this.$q.dialog({message:"Please upload AR and OAT file", title:"Incomplete File"})
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
                  concate : `${m.Asset} ${m['SNo.']}`,
                  concate_or : `${m['Or. asset']} ${m['SNo.']}`
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
          let tanggal_sto = this.dataOAT.find(f=>f.ba === m.bus_a) ? this.dataOAT.find(f=>f.ba === m.bus_a).tanggal_scan:''
          let mutasi = this.dataOAT.find(f=>f.asset_no === m.concate_or) ? 'MUTASI' : 'CEK'
          return {
            ...m,
            find_on_oat: this.dataOAT.find(f=>f.asset_no === m.concate) ? '' : 'N/A',
            tanggal_sto : tanggal_sto,
            check : m.cap_date > tanggal_sto ? 'BELI' : mutasi
          }
        })
        for (const [key, value] of Object.entries(this.resultARStep2[1])) {
            this.resultARStep2_fields[key] = key;
        }
        this.resultARStep2 = this.resultARStep2.filter(f=>f.find_on_oat ==='N/A')
        this.resultOATStep2 = this.dataOAT.map(m=>{
          let mutasi_ba = this.dataAR.find(f=>f.concate_or === m.asset_no) ? `MUTASI KE ${this.dataAR.find(f=>f.concate_or === m.asset_no).bus_a}` : 'CEK'
          return {
            ...m,
            find_on_ar: this.dataAR.find(f=>f.concate === m.asset_no) ? '' : 'N/A',
            find_mutasi:  mutasi_ba
          }
        })
        for (const [key, value] of Object.entries(this.resultOATStep2[5])) {
            this.resultOATStep2_fields[key] = key;
        }
        this.resultOATStep2 = this.resultOATStep2.filter(f=>f.find_on_ar ==='N/A')
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
                workbook.Sheets[sheetName], {raw:true}
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