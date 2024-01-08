use calamine::{Reader, open_workbook, Xlsx, DataType};
use chrono::{DateTime, Utc, TimeZone};


const TODO_ID_COL: u32 = 0;  // TODO: 231225 本当は、VBA のコードの中から見るのが良い。ただ、なぜか、a_main.bas が取れずにいる。。
const MAIN_CLASS_COL: u32 = 1;
const SUB_CLASS_COL: u32 = 2;
const START_DATE_COL: u32 = 3;
const END_DATE_COL: u32 = 4;
const CONTENT_COL: u32 = 5;  // INFO: 231225 0 始まりの index であることに注意せよ。
const SATART_EACH_TASK_COL: u32 = 6;

const DATE_IDX: u32 = 0;
const START_TODO_IDX: u32 = 6;


fn main() {
    let path = "./DailyTask.xlsm";
    let sheet_name = "Sheet1";
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");

    if let Ok(range) = workbook.worksheet_range(sheet_name) {
        let max_idx: u32 = range.get_size().0.try_into().unwrap();
        let max_col: u32 = range.get_size().1.try_into().unwrap();

        for idx in START_TODO_IDX..max_idx {
            if let DataType::Float(todo_id) = range.get_value((idx, TODO_ID_COL)).unwrap() {
                let todo_id_ = *todo_id as i64;
                let todo_summary = SummaryTask {
                    todo_id: todo_id_,
                    main_class: range.get_value((idx, MAIN_CLASS_COL)).unwrap().as_string(),
                    sub_class: range.get_value((idx, SUB_CLASS_COL)).unwrap().as_string(),
                    start_date: range.get_value((idx, START_DATE_COL)).unwrap().as_i64(),
                    end_date: range.get_value((idx, END_DATE_COL)).unwrap().as_i64(),
                    content: range.get_value((idx, CONTENT_COL)).unwrap().as_string(),
                };
                println!("{:?}", todo_summary);

                for col in SATART_EACH_TASK_COL..max_col {
                    if let DataType::DateTime(value) = range.get_value((DATE_IDX, col)).unwrap() {
                        let date_time = Utc.timestamp_millis_opt(((*value as i64) * 86400 * 1000) - 2209161600000);
                        let date_time = date_time.unwrap().format("%Y-%m-%d").to_string();
                        let each_task = EachTask {
                            todo_id: todo_id_,
                            date: date_time,
                            content: range.get_value((idx, col)).unwrap().as_string(),
                        };
                        if each_task.content != None {
                            println!("{:?}", each_task);
                        }
                    }
                }
            }
        }
        
    } else {
        println!("No Sheet of '{}' ...", sheet_name);
    }
}

#[derive(Debug)]
struct SummaryTask {
    todo_id: i64,
    main_class: Option<String>,
    sub_class: Option<String>,
    start_date: Option<i64>,
    end_date: Option<i64>,
    content: Option<String>,
}

#[derive(Debug)]
struct EachTask {
    todo_id: i64,
    date: String,
    content: Option<String>,
}