CREATE TABLE IF NOT EXISTS user_professions (
    user_id       INTEGER NOT NULL,
    profession_id INTEGER NOT NULL,
    learned_at    INTEGER NOT NULL,
    PRIMARY KEY (user_id, profession_id)
);
CREATE INDEX IF NOT EXISTS idx_user_professions_user ON user_professions(user_id);
ALTER TABLE user ADD COLUMN profession_forgot_count INTEGER NOT NULL DEFAULT 0;
