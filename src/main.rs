use calamine::{Reader, open_workbook, Xlsx, DataType};
use chrono::{DateTime, Utc, TimeZone};
use sqlx::{migrate::MigrateDatabase, FromRow, Row, Sqlite, SqlitePool};
use dotenv;
use std::env;


const TODO_ID_COL: u32 = 0;  // TODO: 231225 本当は、VBA のコードの中から見るのが良い。ただ、なぜか、a_main.bas が取れずにいる。。
const MAIN_CLASS_COL: u32 = 1;
const SUB_CLASS_COL: u32 = 2;
const START_DATE_COL: u32 = 3;
const END_DATE_COL: u32 = 4;
const CONTENT_COL: u32 = 5;  // INFO: 231225 0 始まりの index であることに注意せよ。
const SATART_EACH_TASK_COL: u32 = 6;

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

        for idx in START_TODO_IDX..max_idx {
            if let DataType::Float(todo_id) = range.get_value((idx, TODO_ID_COL)).unwrap() {
                let todo_summary = SummaryTask {
                    todo_id: *todo_id as i64,
                    main_class: range.get_value((idx, MAIN_CLASS_COL)).unwrap().as_string(),
                    sub_class: range.get_value((idx, SUB_CLASS_COL)).unwrap().as_string(),
                    start_date: cast_excel_date_to_i64(range.get_value((idx, START_DATE_COL))),
                    end_date: cast_excel_date_to_i64(range.get_value((idx, END_DATE_COL))),
                    content: range.get_value((idx, CONTENT_COL)).unwrap().as_string(),
                };
                todo_summary.upsert().await;


                // for col in SATART_EACH_TASK_COL..max_col {
                //     if let DataType::DateTime(value) = range.get_value((DATE_IDX, col)).unwrap() {
                //         let date_time = Utc.timestamp_millis_opt(((*value as i64) * 86400 * 1000) - 2209161600000);
                //         let date_time = date_time.unwrap().format("%Y-%m-%d").to_string();
                //         let each_task = EachTask {
                //             todo_id: todo_id_,
                //             date: date_time,
                //             content: range.get_value((idx, col)).unwrap().as_string(),
                //         };
                //         if each_task.content != None {
                //             println!("{:?}", each_task);
                //         }
                //     }
                // }
            }
        }
        
    } else {
        println!("No Sheet of '{}' ...", sheet_name);
    }
}


/// calamine で取得したDateTime を、integer に変換する関数。
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


/// sqlx で、データベースプールオブジェクトを取得する関数。<br>
/// "DATABASE_URL" は存在する前提で、非存在時は panic となる。(sqlx でマクロメインで実装するため、別で .db ファイルを作るものとする。)
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

// TODO: 240109 単機能で入力できるかのテストを実装する？(テスト用のデータベースを準備するのか？)
impl SummaryTask {
    async fn upsert(&self) {
        let db = obtain_db_connection().await.unwrap();
        let temp_result = sqlx::query_as!(SummaryTask, "SELECT * FROM summary WHERE todo_id = ?", self.todo_id).fetch_all(&db).await.unwrap();
        if temp_result.len() == 0 {
            let result = sqlx::query!(
                "INSERT INTO summary (todo_id, main_class, sub_class, start_date, end_date, content) VALUES (?, ?, ?, ?, ?, ?)",
                self.todo_id,
                self.main_class,
                self.sub_class,
                self.start_date,
                self.end_date,
                self.content,
            )
            .execute(&db)
            .await
            .unwrap();
            // println!("Query result: {:?}", result);  // TODO: 240109 この関数全体が、Result 型で返すべきな気もする。

        } else {  // TODO: 240109 1 以上がここに来ることになるが、todo_id はユニークであるはずなので、2 以上が来たらエラーを返すべきな気もする。
            let result = sqlx::query!(
                "UPDATE summary SET todo_id = ?, main_class = ?, sub_class = ?, start_date = ?, end_date = ?, content = ? WHERE todo_id = ?",
                self.todo_id,
                self.main_class,
                self.sub_class,
                self.start_date,
                self.end_date,
                self.content,
                self.todo_id,
            )
            .execute(&db)
            .await
            .unwrap();
            // println!("Query result: {:?}", result);  // TODO: 240109 更新はそう起きないはずなので、注意喚起の意味でログを残すと良いのかも？あと、削除されたかのチェックも入れると良い？
        }
    }
}

#[derive(Debug)]
struct EachTask {
    todo_id: i64,
    date: String,
    content: Option<String>,
}