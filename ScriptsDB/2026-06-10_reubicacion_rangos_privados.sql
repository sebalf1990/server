-- Plan 10.003: relocation of private ID ranges (Heroism collision).
-- Run with the game server STOPPED. Backup database.db first.
--
-- Spell IDs: poison system + Detectar Personajes moved 295-307 -> 400-412 (+105).
-- Item IDs: profession manuals/potions moved 4997-5010 -> 9000-9013 (+4003).
-- Player-persisted data must follow the new IDs.

BEGIN TRANSACTION;

UPDATE spell
SET spell_id = spell_id + 105
WHERE spell_id BETWEEN 295 AND 307;

UPDATE inventory_item
SET item_id = item_id + 4003
WHERE item_id BETWEEN 4997 AND 5010;

UPDATE bank_item
SET item_id = item_id + 4003
WHERE item_id BETWEEN 4997 AND 5010;

COMMIT;
