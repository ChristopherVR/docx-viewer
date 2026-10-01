// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Relationship parsing and part-path resolution live in the shared `@christophervr/ooxml-core/opc`;
// this module keeps the existing import path working.
export {
	getRelationshipId,
	parseRelationships,
	resolvePartPath,
	type Relationship,
} from '@christophervr/ooxml-core/opc';
import { NS } from '@christophervr/ooxml-core/xml';

export const RELATIONSHIP_NS = NS.r;
