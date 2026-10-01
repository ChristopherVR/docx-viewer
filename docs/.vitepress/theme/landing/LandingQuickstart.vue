<script setup lang="ts">
import { withBase } from 'vitepress';
import { ref } from 'vue';

import CodeCard from './code/CodeCard.vue';
import { FRAMEWORKS } from './code/samples';
import { useLandingCopy } from './copy';

const copy = useLandingCopy();
const active = ref(FRAMEWORKS[0]);

function select(id: string): void {
	active.value = FRAMEWORKS.find((f) => f.id === id) ?? FRAMEWORKS[0];
}
</script>

<template>
	<section class="pv-section pv-quickstart">
		<p class="pv-kicker" data-reveal>{{ copy.quickstart.kicker }}</p>
		<h2 class="pv-h2" data-reveal="2">{{ copy.quickstart.title }}</h2>
		<p class="pv-copy" data-reveal="3">{{ copy.quickstart.copy }}</p>
		<div class="pv-quickstart__panel" data-reveal="3">
			<div class="pv-quickstart__tabs" role="tablist">
				<button
					v-for="fw in FRAMEWORKS"
					:key="fw.id"
					type="button"
					role="tab"
					:aria-selected="fw.id === active.id"
					:class="['pv-quickstart__tab', { 'is-active': fw.id === active.id }]"
					@click="select(fw.id)"
				>
					{{ fw.label }}
				</button>
			</div>
			<CodeCard :file="active.file" :badge="active.entry" :code="active.code" />
			<div class="pv-quickstart__build">
				<strong>{{ copy.quickstart.buildTitle }}</strong>
				<p>{{ copy.quickstart.buildCopy }}</p>
				<pre><code>{{ copy.quickstart.buildCommands }}</code></pre>
			</div>
			<a class="pv-link" :href="withBase(active.docsHref)">
				{{ copy.quickstart.docsLabel }}: {{ active.label }} <i>&rarr;</i>
			</a>
		</div>
	</section>
</template>

<style scoped>
.pv-quickstart__panel {
	margin-top: 2.2rem;
	max-width: 54rem;
}

.pv-quickstart__tabs {
	display: flex;
	flex-wrap: wrap;
	gap: 0.35rem;
	margin-bottom: 1rem;
}

.pv-quickstart__tab {
	font-family: var(--pv-mono);
	font-size: 0.72rem;
	letter-spacing: 0.14em;
	text-transform: uppercase;
	color: var(--pv-ink-soft);
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-radius: 3px;
	padding: 0.55em 1em;
	cursor: pointer;
	transition:
		color 0.2s ease,
		border-color 0.2s ease;
}

.pv-quickstart__tab:hover {
	color: var(--pv-ink);
	border-color: var(--pv-ink-soft);
}

.pv-quickstart__tab.is-active {
	color: #fff;
	background: var(--pv-accent);
	border-color: var(--pv-accent);
}

.pv-quickstart__build {
	margin-top: 1.2rem;
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-left: 3px solid var(--pv-accent);
	border-radius: 6px;
	padding: 1.1rem 1.3rem 1.2rem;
}

.pv-quickstart__build strong {
	font-family: var(--pv-mono);
	font-size: 0.68rem;
	letter-spacing: 0.22em;
	text-transform: uppercase;
	color: var(--pv-accent);
}

.pv-quickstart__build p {
	margin: 0.6rem 0 0.8rem;
	font-size: 0.88rem;
	line-height: 1.7;
	color: var(--pv-ink-soft);
}

.pv-quickstart__build pre {
	margin: 0;
	padding: 0.7rem 0.9rem;
	background: #14171b;
	border-radius: 4px;
	overflow-x: auto;
}

.pv-quickstart__build code {
	font-family: var(--pv-mono);
	font-size: 0.78rem;
	line-height: 1.7;
	color: #dfe3e8;
	background: none;
	white-space: pre;
}
</style>
