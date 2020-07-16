import Sheet from './Sheet.vue' // 导入组件
Sheet.install = (Vue) => {
    Vue.component('Sheet', Sheet)
    if (typeof window !== 'undefined' && window.Vue) {
        window.Vue.use(Sheet);
    }
}
export default Sheet