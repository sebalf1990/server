UPDATE account SET email = lower(email);
UPDATE account SET validated = 1;
CREATE UNIQUE INDEX IF NOT EXISTS idx_account_email_lower ON account (lower(email));
