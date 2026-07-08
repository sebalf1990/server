CREATE UNIQUE INDEX IF NOT EXISTS idx_account_email_lower ON account (lower(email));
