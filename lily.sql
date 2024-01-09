CREATE TABLE IF NOT EXISTS summary (
    todo_id    INTEGER PRIMARY KEY,
    main_class TEXT,
    sub_class  TEXT,
    start_date INTEGER,
    end_date   INTEGER,
    content    TEXT
);

CREATE TABLE IF NOT EXISTS content (
    todo_id  INTEGER NOT NULL,
    date     INTEGER NOT NULL,
    content  TEXT,
    PRIMARY KEY (todo_id, date)
);