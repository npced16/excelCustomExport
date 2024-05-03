import { createRouter, createWebHashHistory } from 'vue-router';
import  leageW  from "@/components/leageW.vue";
const routes = [
	{
		path: "/language",
		name: "Webpage",
		component: leageW,
	},
];

const router = createRouter({
	history: createWebHashHistory(process.env.BASE_URL),
	routes,
});

export default router;
