<script setup lang="ts">
import { withBase } from 'vitepress';

import CodeCard from './code/CodeCard.vue';
import { CORE_SAMPLE } from './code/samples';
import { useLandingCopy } from './copy';

const copy = useLandingCopy();
</script>

<template>
	<section class="pv-section pv-status">
		<div class="pv-status__grid">
			<div class="pv-status__copy">
				<p class="pv-kicker" data-reveal>{{ copy.status.kicker }}</p>
				<h2 class="pv-h2" data-reveal="2">{{ copy.status.title }}</h2>
				<p class="pv-copy" data-reveal="3">{{ copy.status.copy }}</p>
				<a class="pv-link" :href="withBase(copy.status.link.href)" data-reveal="3">
					{{ copy.status.link.text }} <i>&rarr;</i>
				</a>
			</div>
			<div class="pv-status__code" data-reveal="3">
				<CodeCard file="load-save.ts" badge="DOCX &middot; DOC" :code="CORE_SAMPLE" />
				<aside class="pv-status__note">
					<strong>{{ copy.status.roadmapTitle }}</strong>
					<p>{{ copy.status.roadmapCopy }}</p>
					<a :href="withBase(copy.status.roadmapLink.href)">
						{{ copy.status.roadmapLink.text }} &rarr;
					</a>
				</aside>
			</div>
		</div>
	</section>
</template>

<style scoped>
.pv-status__grid {
	display: grid;
	grid-template-columns: minmax(0, 1fr) minmax(0, 1.1fr);
	gap: clamp(2rem, 5vw, 5rem);
	align-items: center;
}

.pv-status__copy .pv-copy {
	max-width: 33rem;
}

.pv-status__code {
	display: flex;
	flex-direction: column;
	gap: 1.2rem;
	min-width: 0;
}

.pv-status__note {
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-left: 3px solid var(--pv-accent);
	border-radius: 6px;
	padding: 1.1rem 1.3rem 1.2rem;
}

.pv-status__note strong {
	font-family: var(--pv-mono);
	font-size: 0.68rem;
	letter-spacing: 0.22em;
	text-transform: uppercase;
	color: var(--pv-accent);
}

.pv-status__note p {
	margin: 0.6rem 0 0.7rem;
	font-size: 0.88rem;
	line-height: 1.7;
	color: var(--pv-ink-soft);
}

.pv-status__note a {
	font-family: var(--pv-mono);
	font-size: 0.74rem;
	letter-spacing: 0.08em;
	color: var(--pv-accent);
}

@media (max-width: 900px) {
	.pv-status__grid {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
