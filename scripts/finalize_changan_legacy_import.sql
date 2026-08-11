BEGIN;

UPDATE source_chunks
SET correction_status = COALESCE(payload->>'correction_status', correction_status),
    updated_at = COALESCE(updated_at, now())
WHERE data_classification = 'research_pilot';

UPDATE sources
SET confidence = COALESCE((payload->>'confidence')::double precision, confidence),
    updated_at = COALESCE(updated_at, now())
WHERE data_classification = 'research_pilot';

COMMIT;
