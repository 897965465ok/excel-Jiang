import Vue from 'vue'
import App from './App.vue'
import Sheet from './lib/index.js'
Vue.use(Sheet)
new Vue({
  el: '#app',
  render: h => h(App)
})
