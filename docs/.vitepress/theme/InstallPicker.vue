<script setup lang="ts">
import { ref } from 'vue';
const entries = ['react', 'vue', 'angular', 'vanilla', 'svelte'];
const active = ref('react');
const feedback = ref('Copy');
const command = 'npm install @christophervr/docx-viewer';
async function copy() {
	try {
		await navigator.clipboard.writeText(command);
		feedback.value = 'Copied';
	} catch {
		feedback.value = 'Select and copy the command below';
	}
}
</script>
<template>
	<div class="dv-install">
		<span>ONE PACKAGE · COMING TO NPM</span>
		<div class="dv-tabs" role="group" aria-label="Installation binding">
			<button
				v-for="entry in entries"
				:key="entry"
				:aria-pressed="entry === active"
				@click="
					active = entry;
					feedback = 'Copy';
				"
			>
				{{ entry === 'vanilla' ? 'Vanilla JS' : entry.charAt(0).toUpperCase() + entry.slice(1) }}
			</button>
		</div>
		<div class="dv-install-command">
			<code>{{ command }}</code
			><button @click="copy">{{ feedback }}</button>
		</div>
		<p>
			Import from <code>@christophervr/docx-viewer/{{ active }}</code>
		</p>
	</div>
</template>
<style scoped>
.dv-install-command {
	display: flex;
	gap: 1rem;
	align-items: center;
	margin-top: 0.9rem;
	padding: 0.8rem;
	border: 1px solid var(--dv-line);
	border-radius: 3px;
	overflow-x: auto;
}
.dv-install-command code {
	white-space: nowrap;
	font-size: 0.75rem;
}
.dv-install-command button {
	margin-left: auto;
	color: var(--dv-accent);
	font: 500 0.7rem var(--docx-mono);
}
p {
	margin-top: 0.7rem;
	font-size: 0.7rem;
	color: var(--dv-muted);
	overflow-wrap: anywhere;
}
</style>
