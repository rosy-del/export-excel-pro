//import vue文件的name from 路径 
import FirstComponent from "../packages/components/exportExcelStyle.js"; 

const components = [
    FirstComponent
];

const install = function (Vue) {
    if (install.installed) return;
    install.installed = true;
    components.forEach(component => Vue.component(component.name, component));
};

if (typeof window !== 'undefined' && window.Vue) {
    install(window.Vue);
}

// 暴露安装方法和组件
const ChatRobotPlugin = {
    install,
    FirstComponent
};

export default ChatRobotPlugin;
export { FirstComponent };
