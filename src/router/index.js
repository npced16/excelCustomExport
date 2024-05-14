import { createRouter, createWebHashHistory } from 'vue-router';
import languageDownload from "@/components/languageDownload.vue";
const routes = [
	{
		path: "/language",
		name: "Webpage",
		component: languageDownload,
	},
];

const router = createRouter({
	history: createWebHashHistory(process.env.BASE_URL),
	routes,
});

export default router;
