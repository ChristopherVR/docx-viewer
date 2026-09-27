import DefaultTheme from 'vitepress/theme';
import type { Theme } from 'vitepress';
import './custom.css';
import LandingHome from './LandingHome.vue';

export default {
	extends: DefaultTheme,
	enhanceApp({ app }) {
		app.component('LandingHome', LandingHome);
	},
} satisfies Theme;
