CREATE TABLE IF NOT EXISTS summary (
    todo_id    INTEGER PRIMARY KEY,
    main_class TEXT,
    sub_class  TEXT,
    start_date TEXT,
    end_date   TEXT,
    content    TEXT
);

CREATE TABLE IF NOT EXISTS content (
    todo_id  INTEGER PRIMARY KEY,
    date     TEXT NOT NULL,
    content  TEXT
);