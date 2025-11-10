-- schema.sql
CREATE INDEX IF NOT EXISTS idx_equity_index_constituents_lookup
ON equity_index_constituents (index_name, accord_code, date);
