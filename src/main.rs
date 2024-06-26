use calamine::{Reader, open_workbook, Xlsx, DataType};
use sqlx::{FromRow, Sqlite, SqlitePool};
use dotenv;
use std::env;
use indicatif::ProgressIterator;


const TODO_ID_COL: u32 = 0;
const MAIN_CLASS_COL: u32 = 1;
const SUB_CLASS_COL: u32 = 2;
const START_DATE_COL: u32 = 3;
const END_DATE_COL: u32 = 4;
const CONTENT_COL: u32 = 5;
const START_EACH_TASK_COL: u32 = 6;

const DATE_IDX: u32 = 0;
const START_TODO_IDX: u32 = 6;


#[tokio::main]
async fn main() {
    let path = "./DailyTask.xlsm";  // TODO: 240109 walkdir で複数ファイルで実行できるようにせよ。
    let sheet_name = "Sheet1";
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");

    if let Ok(range) = workbook.worksheet_range(sheet_name) {
        let max_idx: u32 = range.get_size().0.try_into().unwrap();
        let max_col: u32 = range.get_size().1.try_into().unwrap();
        let db_conn = obtain_db_connection().await.unwrap();

        for idx in (START_TODO_IDX..max_idx).progress() {  // INFO: 240109 .progress() は、indicatif のプログレスバー出力。
            if let DataType::Float(todo_id) = range.get_value((idx, TODO_ID_COL)).unwrap() {
                let todo_summary = SummaryTask {
                    todo_id: *todo_id as i64,
                    main_class: range.get_value((idx, MAIN_CLASS_COL)).unwrap().as_string(),
                    sub_class: range.get_value((idx, SUB_CLASS_COL)).unwrap().as_string(),
                    start_date: cast_excel_date_to_i64(range.get_value((idx, START_DATE_COL))),
                    end_date: cast_excel_date_to_i64(range.get_value((idx, END_DATE_COL))),
                    content: change_valid_crlf_word_of_excel(range.get_value((idx, CONTENT_COL))),
                };
                let _result = todo_summary.upsert(&db_conn).await;

                for col in START_EACH_TASK_COL..max_col {
                    if let DataType::DateTime(value) = range.get_value((DATE_IDX, col)).unwrap() {
                        let each_task = EachTask {
                            todo_id: *todo_id as i64,
                            date: *value as i64,
                            content: range.get_value((idx, col)).unwrap().as_string(),
                        };
                        let _result = each_task.upsert(&db_conn).await;
                    }
                }
            }
        }
        
    } else {
        panic!("FatalError: No Sheet of '{}' ...", sheet_name);
    }
}


/// calamine で取得した DateTime を、integer に変換する関数。
fn cast_excel_date_to_i64(cell: Option<&DataType>) -> Option<i64> {
    let cell = cell.unwrap();
    match cell {
        DataType::DateTime(date_time) => {
            return Some(*date_time as i64);
        }
        _ => {
            return None;
        }
    }
}


/// エクセルの不具合なのか、_x000D_ が改行コードして保持されているケースがあるようで、その文字列を変更する関数。
fn change_valid_crlf_word_of_excel(cell: Option<&DataType>) -> Option<String> {
    let cell = cell.unwrap();
    match cell {
        DataType::String(value) => {
            return Some(value.replace("_x000D_", "").clone());
        }
        _ => {
            return None;
        }
    }
}


/// sqlx で、データベースプールオブジェクトを取得する関数。<br>
/// "DATABASE_URL" は存在する前提で、非存在時は panic となる。(sqlx でマクロメインで実装するため、事前に別で .db ファイルを作るものとする。)
async fn obtain_db_connection() -> sqlx::Result<sqlx::Pool<Sqlite>> {
    dotenv::dotenv().expect("Failed to read .env file");  // INFO: 240109 dotenv::from_filename だと、sqlx のマクロがうまく実行できていないみたいだったので、.env ファイルを対象とした。
    let db_url = env::var("DATABASE_URL").expect("DATABASE_URL must be set");
    Ok(SqlitePool::connect(&db_url).await?)
}

#[derive(Clone, FromRow, Debug)]
struct SummaryTask {
    todo_id: i64,
    main_class: Option<String>,
    sub_class: Option<String>,
    start_date: Option<i64>,
    end_date: Option<i64>,
    content: Option<String>,
}


impl SummaryTask {
    async fn upsert(&self, db_conn: &sqlx::Pool<Sqlite>) -> Result<(), sqlx::Error> {
        let temp_result = sqlx::query_as!(SummaryTask, "SELECT * FROM summary WHERE todo_id = ?", self.todo_id)
        .fetch_all(db_conn)
        .await?;//.unwrap();

        match temp_result.len() {
            0 => {
                let _result = sqlx::query!(
                    "INSERT INTO summary (todo_id, main_class, sub_class, start_date, end_date, content) VALUES (?, ?, ?, ?, ?, ?)",
                    self.todo_id,
                    self.main_class,
                    self.sub_class,
                    self.start_date,
                    self.end_date,
                    self.content,
                )
                .execute(db_conn)
                .await?;
            },
            1 => {
                let _result = sqlx::query!(
                    "UPDATE summary SET todo_id = ?, main_class = ?, sub_class = ?, start_date = ?, end_date = ?, content = ? WHERE todo_id = ?",
                    self.todo_id,
                    self.main_class,
                    self.sub_class,
                    self.start_date,
                    self.end_date,
                    self.content,
                    self.todo_id,
                )
                .execute(db_conn)
                .await?;
            },
            _ => {
                panic!("FetalError: todo_id must be unique ...???");
            }
        }
        Ok(())
    }
}

#[derive(Debug)]
struct EachTask {
    todo_id: i64,
    date: i64,
    content: Option<String>,
}

impl EachTask {
    async fn upsert(&self, db_conn: &sqlx::Pool<Sqlite>) -> Result<(), sqlx::Error> {
        let temp_result = sqlx::query_as!(EachTask, "SELECT * FROM content WHERE todo_id = ? AND date = ?", self.todo_id, self.date)
        .fetch_all(db_conn)
        .await?;

        match temp_result.len() {
            0 => {
                if let Some(content_val) = &self.content {
                    let _result = sqlx::query!(
                        "INSERT INTO content (todo_id, date, content) VALUES (?, ?, ?)",
                        self.todo_id,
                        self.date,
                        content_val,
                    )
                    .execute(db_conn)
                    .await?;
                }
            },
            1 => {
                match &self.content {
                    Some(content_val) => {
                        let _result = sqlx::query!(
                            "UPDATE content SET todo_id = ?, date = ?, content = ? WHERE todo_id = ? AND date = ?",
                            self.todo_id,
                            self.date,
                            content_val,
                            self.todo_id,
                            self.date,
                        )
                        .execute(db_conn)
                        .await?;
                    },
                    None => {
                        let _result = sqlx::query!("DELETE FROM content WHERE todo_id = ? AND date = ?", self.todo_id, self.date)
                        .execute(db_conn)
                        .await?;
                    }
                }
            },
            _ => {
                panic!("FetalError: todo_id and date must be unique ...???");
            }
        }

        Ok(())
    }
}