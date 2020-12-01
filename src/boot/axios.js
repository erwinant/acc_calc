import Vue from 'vue'
import axios from 'axios'
import JsonExcel from "vue-json-excel";

Vue.component("downloadExcel", JsonExcel);

Vue.prototype.$axios = axios
