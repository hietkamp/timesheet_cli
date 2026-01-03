use clap::{Parser, Subcommand};
use inquire::{Confirm, CustomType, Select, Text};
use prettytable::{format, Cell, Row, Table};
use rusqlite::{params, Connection, OptionalExtension};
use chrono::{Datelike, Local, NaiveDate, Weekday};
use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, Formula, Image, Workbook, column_number_to_name
};
use std::{collections::{BTreeMap, HashMap}, error::Error};
use dotenv::dotenv;
use std::env;

// --- CLI Structure ---
#[derive(Parser)]
#[command(name = "Timesheet App")]
#[command(about = "Track your daily hours per project per week", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    Template,
    Log,
    Month,
    Export,
}
const FONT_NAME: &str = "Verdana";

// --- Data Structs ---
#[derive(Debug, Clone)]
struct Template {
    id: Option<i32>,
    project: String,
    mon: f64, tue: f64, wed: f64, thu: f64, fri: f64, sat: f64, sun: f64,
}

#[derive(Debug, Clone)]
struct Entry {
    id: Option<i32>,
    project: String,
    mon: f64, tue: f64, wed: f64, thu: f64, fri: f64, sat: f64, sun: f64,
}

impl Template {
    fn total(&self) -> f64 {
        self.mon + self.tue + self.wed + self.thu + self.fri + self.sat + self.sun
    }
}

impl Entry {
    fn total(&self) -> f64 {
        self.mon + self.tue + self.wed + self.thu + self.fri + self.sat + self.sun
    }
}

// Use Box<dyn Error> to handle errors from both Sqlite and Xlsxwriter
fn main() -> Result<(), Box<dyn Error>> {
    dotenv().ok(); // Reads the .env file
    let conn = Connection::open("timesheet.db")?;
    init_db(&conn)?;

    let cli = Cli::parse();

    match cli.command {
        Commands::Template => handle_templates(&conn)?,
        Commands::Log => handle_log(&conn)?,
        Commands::Month => handle_month(&conn)?,
        // Updated to pass connection
        Commands::Export => export_timesheet(&conn)?, 
    }

    Ok(())
}

// --- Database Init ---
fn init_db(conn: &Connection) -> Result<(), Box<dyn Error>> {
    conn.execute(
        "CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY,
            project TEXT NOT NULL UNIQUE,
            mon REAL DEFAULT 0, tue REAL DEFAULT 0, wed REAL DEFAULT 0, 
            thu REAL DEFAULT 0, fri REAL DEFAULT 0, sat REAL DEFAULT 0, sun REAL DEFAULT 0
        )",
        [],
    )?;
    conn.execute(
        "CREATE TABLE IF NOT EXISTS timesheets (
            id INTEGER PRIMARY KEY,
            week TEXT NOT NULL,
            project TEXT NOT NULL,
            mon REAL DEFAULT 0, tue REAL DEFAULT 0, wed REAL DEFAULT 0, 
            thu REAL DEFAULT 0, fri REAL DEFAULT 0, sat REAL DEFAULT 0, sun REAL DEFAULT 0,
            UNIQUE(week, project)
        )",
        [],
    )?;
    Ok(())
}

// --- Function 1: Templates ---
fn handle_templates(conn: &Connection) -> Result<(), Box<dyn Error>> {
    loop {
        println!("\n--- Template Management (Daily Defaults) ---");
        
        let mut stmt = conn.prepare("SELECT id, project, mon, tue, wed, thu, fri, sat, sun FROM templates")?;
        let templates_iter = stmt.query_map([], |row| {
            Ok(Template {
                id: Some(row.get(0)?),
                project: row.get(1)?,
                mon: row.get(2)?, tue: row.get(3)?, wed: row.get(4)?,
                thu: row.get(5)?, fri: row.get(6)?, sat: row.get(7)?, sun: row.get(8)?,
            })
        })?;

        let mut templates: Vec<Template> = Vec::new();
        let mut table = Table::new();
        table.set_format(*format::consts::FORMAT_NO_BORDER_LINE_SEPARATOR);
        table.set_titles(Row::new(vec![
            Cell::new("Project"), Cell::new("M"), Cell::new("T"), Cell::new("W"), 
            Cell::new("T"), Cell::new("F"), Cell::new("S"), Cell::new("S"), Cell::new("TOTAL")
        ]));
        
        let (mut sum_m, mut sum_tu, mut sum_w, mut sum_th, mut sum_f, mut sum_sa, mut sum_su, mut week_total) = (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
        
        for t in templates_iter {
            let t = t?;
            templates.push(t.clone()); // Store for selection later
            let row_total = t.total();
            
            sum_m += t.mon; sum_tu += t.tue; sum_w += t.wed; sum_th += t.thu; sum_f += t.fri; sum_sa += t.sat; sum_su += t.sun;
            week_total += row_total;
            table.add_row(Row::new(vec![
                Cell::new(&t.project),
                Cell::new(&format_hours(t.mon)), 
                Cell::new(&format_hours(t.tue)), 
                Cell::new(&format_hours(t.wed)),
                Cell::new(&format_hours(t.thu)), 
                Cell::new(&format_hours(t.fri)), 
                Cell::new(&format_hours(t.sat)), 
                Cell::new(&format_hours(t.sun)), 
                Cell::new(&format_hours(row_total)).style_spec("b")
            ]));
        }

        // --- THE DAY TOTAL ROW ---
        table.add_row(Row::new(vec![
            Cell::new("TOTAL").style_spec("b"),
            Cell::new(&format_hours(sum_m)).style_spec("b"),
            Cell::new(&format_hours(sum_tu)).style_spec("b"),
            Cell::new(&format_hours(sum_w)).style_spec("b"),
            Cell::new(&format_hours(sum_th)).style_spec("b"),
            Cell::new(&format_hours(sum_f)).style_spec("b"),
            Cell::new(&format_hours(sum_sa)).style_spec("b"),
            Cell::new(&format_hours(sum_su)).style_spec("b"),
            Cell::new(&format_hours(week_total)).style_spec("bub"), // Bold Underline Bold
        ]));

        table.printstd();

        let options = vec!["Create New", "Edit Existing", "Delete", "Exit"];
        let choice = Select::new("Action:", options).prompt();

        match choice {
            Ok("Create New") => {
                let project = Text::new("Project Name:").prompt().unwrap_or_default();
                if project.is_empty() { continue; }
                
                // Helper to ask for all days
                let (m, tu, w, th, f, sa, su) = prompt_week_hours(0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0);

                conn.execute(
                    "INSERT INTO templates (project, mon, tue, wed, thu, fri, sat, sun) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8)", 
                    params![project, m, tu, w, th, f, sa, su]
                ).map_err(|e| println!("Error: {}", e)).ok();
            },
            Ok("Edit Existing") => {
                if templates.is_empty() { continue; }
                let selection = Select::new("Select Project:", templates.clone()).prompt();
                if let Ok(tmpl) = selection {
                    // Pre-fill with current values
                    let (m, tu, w, th, f, sa, su) = prompt_week_hours(tmpl.mon, tmpl.tue, tmpl.wed, tmpl.thu, tmpl.fri, tmpl.sat, tmpl.sun);
                    
                    conn.execute(
                        "UPDATE templates SET mon=?1, tue=?2, wed=?3, thu=?4, fri=?5, sat=?6, sun=?7 WHERE id=?8", 
                        params![m, tu, w, th, f, sa, su, tmpl.id]
                    )?;
                }
            },
            Ok("Delete") => {
                if let Ok(tmpl) = Select::new("Select Project:", templates.clone()).prompt() {
                     if Confirm::new("Are you sure?").prompt().unwrap_or(false) {
                        conn.execute("DELETE FROM templates WHERE id = ?1", params![tmpl.id])?;
                     }
                }
            },
            _ => break,
        }
    }
    Ok(())
}

// --- Function 2: Timesheets ---
fn handle_log(conn: &Connection) -> Result<(), Box<dyn Error>> {
    let current_date = Local::now();
    let default_week = format!("{}-W{:02}", current_date.year(), current_date.iso_week().week()+1);
    
    let week = Text::new("Enter Week (YYYY-W##):")
        .with_default(&default_week)
        .prompt()
        .unwrap_or(default_week);

    loop {
        // Load entries
        let mut stmt = conn.prepare("SELECT id, project, mon, tue, wed, thu, fri, sat, sun FROM timesheets WHERE week = ?1")?;
        let entries_iter = stmt.query_map(params![week], |row| {
            Ok(Entry {
                id: Some(row.get(0)?),
                project: row.get(1)?,
                mon: row.get(2)?, tue: row.get(3)?, wed: row.get(4)?,
                thu: row.get(5)?, fri: row.get(6)?, sat: row.get(7)?, sun: row.get(8)?,
            })
        })?;

        let mut entries = Vec::new();
        for e in entries_iter { entries.push(e?); }

        if entries.is_empty() {
             println!("No entries found for {}.", week);
             if Confirm::new("Load defaults from Templates?").prompt().unwrap_or(false) {
                conn.execute(
                    "INSERT INTO timesheets (week, project, mon, tue, wed, thu, fri, sat, sun) 
                     SELECT ?1, project, mon, tue, wed, thu, fri, sat, sun FROM templates", 
                    params![week]
                )?;
                continue; 
             }
        }

        // Display Table
        println!("\n--- Timesheet: {} ---", week);
        let mut table = Table::new();
        table.set_format(*format::consts::FORMAT_NO_BORDER_LINE_SEPARATOR);
        table.set_titles(Row::new(vec![
            Cell::new("Project"), 
            Cell::new("Mon"), Cell::new("Tue"), Cell::new("Wed"), Cell::new("Thu"), 
            Cell::new("Fri"), Cell::new("Sat"), Cell::new("Sun"), Cell::new("TOTAL")
        ]));

        let (mut sum_m, mut sum_tu, mut sum_w, mut sum_th, mut sum_f, mut sum_sa, mut sum_su, mut week_total) = (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
        
        for entry in &entries {
            let row_total = entry.total();
            sum_m += entry.mon; sum_tu += entry.tue; sum_w += entry.wed; sum_th += entry.thu; sum_f += entry.fri; sum_sa += entry.sat; sum_su += entry.sun;
            week_total += row_total;
            table.add_row(Row::new(vec![
                Cell::new(&entry.project),
                Cell::new(&format_hours(entry.mon)), 
                Cell::new(&format_hours(entry.tue)), 
                Cell::new(&format_hours(entry.wed)),
                Cell::new(&format_hours(entry.thu)), 
                Cell::new(&format_hours(entry.fri)), 
                Cell::new(&format_hours(entry.sat)), 
                Cell::new(&format_hours(entry.sun)), 
                Cell::new(&format_hours(row_total)).style_spec("b")
            ]));
        } 
        
        // --- THE DAY TOTAL ROW ---
        table.add_row(Row::new(vec![
            Cell::new("TOTAL").style_spec("b"),
            Cell::new(&format_hours(sum_m)).style_spec("b"),
            Cell::new(&format_hours(sum_tu)).style_spec("b"),
            Cell::new(&format_hours(sum_w)).style_spec("b"),
            Cell::new(&format_hours(sum_th)).style_spec("b"),
            Cell::new(&format_hours(sum_f)).style_spec("b"),
            Cell::new(&format_hours(sum_sa)).style_spec("b"),
            Cell::new(&format_hours(sum_su)).style_spec("b"),
            Cell::new(&format_hours(week_total)).style_spec("bub"), // Bold Underline Bold
        ]));
        
        table.printstd();

        // Menu
        let action = Select::new("Action:", vec!["Edit Day", "Add Project", "Remove Project", "Exit"]).prompt();

        match action {
            Ok("Edit Day") => {
                if entries.is_empty() { continue; }
                // 1. Select Project
                let entry = Select::new("Select Project:", entries.clone()).prompt();
                if let Ok(e) = entry {
                    // 2. Select Day
                    let days = vec!["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
                    let day_choice = Select::new("Select Day:", days).prompt();
                    
                    if let Ok(d) = day_choice {
                        let current_val = match d { "Mon"=>e.mon, "Tue"=>e.tue, "Wed"=>e.wed, "Thu"=>e.thu, "Fri"=>e.fri, "Sat"=>e.sat, _=>e.sun };
                        let new_val = CustomType::<f64>::new(&format!("Hours for {}:", d)).with_default(current_val).prompt().unwrap_or(current_val);
                        
                        let col_name = match d { "Mon"=>"mon", "Tue"=>"tue", "Wed"=>"wed", "Thu"=>"thu", "Fri"=>"fri", "Sat"=>"sat", _=>"sun" };
                        let sql = format!("UPDATE timesheets SET {} = ?1 WHERE id = ?2", col_name);
                        
                        conn.execute(&sql, params![new_val, e.id])?;
                    }
                }
            },
            Ok("Add Project") => {
                let project = Text::new("Project Name:").prompt().unwrap_or_default();
                if !project.is_empty() {
                    conn.execute(
                        "INSERT INTO timesheets (week, project) VALUES (?1, ?2)", 
                        params![week, project]
                    ).ok(); 
                }
            },
            Ok("Remove Project") => {
                 if let Ok(entry) = Select::new("Remove:", entries).prompt() {
                     conn.execute("DELETE FROM timesheets WHERE id = ?1", params![entry.id])?;
                 }
            },
            _ => break,
        }
    }
    Ok(())
}

// Helper to prompt for 7 days quickly
fn prompt_week_hours(m:f64, tu:f64, w:f64, th:f64, f:f64, sa:f64, su:f64) -> (f64, f64, f64, f64, f64, f64, f64) {
    println!("Enter hours for each day (Press Enter to keep default):");
    let nm = CustomType::<f64>::new("Mon:").with_default(m).prompt().unwrap_or(m);
    let ntu = CustomType::<f64>::new("Tue:").with_default(tu).prompt().unwrap_or(tu);
    let nw = CustomType::<f64>::new("Wed:").with_default(w).prompt().unwrap_or(w);
    let nth = CustomType::<f64>::new("Thu:").with_default(th).prompt().unwrap_or(th);
    let nf = CustomType::<f64>::new("Fri:").with_default(f).prompt().unwrap_or(f);
    let nsa = CustomType::<f64>::new("Sat:").with_default(sa).prompt().unwrap_or(sa);
    let nsu = CustomType::<f64>::new("Sun:").with_default(su).prompt().unwrap_or(su);
    (nm, ntu, nw, nth, nf, nsa, nsu)
}

// Helpers for Display
impl std::fmt::Display for Template {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.project)
    }
}
impl std::fmt::Display for Entry {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.project)
    }
}

fn format_hours(h: f64) -> String {
    if h == 0.0 {
        "".to_string()
    } else {
        h.to_string()
    }
}

// --- Function 4: Monthly Overview (Matrix: Projects vs Days) ---
fn handle_month(conn: &Connection) -> Result<(), Box<dyn Error>> {
    // 1. Defaults
    let now = Local::now();
    let default_year = now.year();
    let default_month = if now.month() == 1 { 12 } else { now.month() - 1 };

    println!("\n--- Monthly Overview (Matrix View) ---");

    // 2. Input
    let selected_year = CustomType::<i32>::new("Year:")
        .with_default(default_year)
        .prompt()?;
    let selected_month = CustomType::<u32>::new("Month (1-12):")
        .with_default(default_month)
        .prompt()?;

    // 3. Calculate Days in Month dynamically
    let next_month = if selected_month == 12 { 1 } else { selected_month + 1 };
    let next_year_val = if selected_month == 12 { selected_year + 1 } else { selected_year };
    
    let days_in_month = NaiveDate::from_ymd_opt(next_year_val, next_month, 1)
        .ok_or("Invalid Date Calculation")?
        .pred_opt()
        .ok_or("Invalid Date Predecessor")?
        .day();

    // 4. Fetch & Aggregate Data
    let mut project_rows: BTreeMap<String, HashMap<u32, f64>> = BTreeMap::new();
    let mut col_totals: HashMap<u32, f64> = HashMap::new();
    let mut grand_total = 0.0;

    let mut stmt = conn.prepare("SELECT week, project, mon, tue, wed, thu, fri, sat, sun FROM timesheets")?;
    let rows = stmt.query_map([], |row| {
         Ok((
            row.get::<_, String>(0)?, // week
            row.get::<_, String>(1)?, // project
            [
                row.get::<_, f64>(2)?, row.get::<_, f64>(3)?, row.get::<_, f64>(4)?, 
                row.get::<_, f64>(5)?, row.get::<_, f64>(6)?, row.get::<_, f64>(7)?, row.get::<_, f64>(8)?
            ] 
        ))
    })?;

    for r in rows {
        let (week_str, project, hours) = r?;
        let parts: Vec<&str> = week_str.split("-W").collect();
        if parts.len() != 2 { continue; }
        
        let w_year: i32 = parts[0].parse().unwrap_or_default();
        let w_num: u32 = parts[1].parse().unwrap_or_default();

        for (day_idx, &h) in hours.iter().enumerate() {
            if h == 0.0 { continue; }

            let weekday = match day_idx {
                0 => Weekday::Mon, 1 => Weekday::Tue, 2 => Weekday::Wed,
                3 => Weekday::Thu, 4 => Weekday::Fri, 5 => Weekday::Sat, _ => Weekday::Sun
            };

            if let Some(date) = NaiveDate::from_isoywd_opt(w_year, w_num, weekday) {
                if date.year() == selected_year && date.month() == selected_month {
                    let d = date.day();
                    *project_rows.entry(project.clone()).or_default().entry(d).or_insert(0.0) += h;
                    *col_totals.entry(d).or_insert(0.0) += h;
                    grand_total += h;
                }
            }
        }
    }

    if project_rows.is_empty() {
        println!("No data found for {}/{}.", selected_month, selected_year);
        return Ok(());
    }

    // 5. Construct Table
    let mut table = Table::new();
    table.set_format(*format::consts::FORMAT_NO_BORDER_LINE_SEPARATOR);

    // --- Header Row: Project | Mon 01 | Tue 02 | ... | TOT ---
    let mut header_cells = vec![Cell::new("Project").style_spec("b")];
    
    for d in 1..=days_in_month {
        // Create date object to get the day name
        if let Some(date) = NaiveDate::from_ymd_opt(selected_year, selected_month, d) {
            let day_name = date.format("%a").to_string(); // e.g., "Mon", "Tue"
            // Format: Name on top, Number below (e.g., "Mon\n01")
            // Style "bc" = Bold + Center alignment
            header_cells.push(Cell::new(&format!("{}\n{:02}", day_name, d)).style_spec("bc"));
        } else {
             header_cells.push(Cell::new(&format!("{:02}", d)).style_spec("bc"));
        }
    }
    header_cells.push(Cell::new("TOT").style_spec("b"));
    table.set_titles(Row::new(header_cells));

    // --- Project Rows ---
    for (proj, days_map) in project_rows {
        let mut row_cells = vec![Cell::new(&proj)];
        let mut row_sum = 0.0;

        for d in 1..=days_in_month {
            let val = days_map.get(&d).unwrap_or(&0.0);
            row_sum += val;
            
            if *val > 0.0 {
                // Align numbers to center as well to match headers
                row_cells.push(Cell::new(&format_hours(*val)).style_spec("c"));
            } else {
                row_cells.push(Cell::new("")); 
            }
        }
        row_cells.push(Cell::new(&format_hours(row_sum)).style_spec("b"));
        table.add_row(Row::new(row_cells));
    }

    // --- Bottom Total Row ---
    let mut footer_cells = vec![Cell::new("TOTAL").style_spec("b")];
    for d in 1..=days_in_month {
        let val = col_totals.get(&d).unwrap_or(&0.0);
        if *val > 0.0 {
            footer_cells.push(Cell::new(&format_hours(*val)).style_spec("bc"));
        } else {
             footer_cells.push(Cell::new(""));
        }
    }
    footer_cells.push(Cell::new(&format_hours(grand_total)).style_spec("bub"));
    table.add_row(Row::new(footer_cells));

    println!("\nReport: {}/{}", selected_month, selected_year);
    table.printstd();

    Ok(())
}

// --- Function 3: Export Timesheet to Excel ---
fn export_timesheet(conn: &Connection) -> Result<(), Box<dyn Error>> {
    // 1. Get Distinct Projects for Selection
    let mut stmt = conn.prepare("SELECT DISTINCT project FROM timesheets ORDER BY project")?;
    let projects_iter = stmt.query_map([], |row| row.get::<_, String>(0))?;
    let mut projects = Vec::new();
    for p in projects_iter { projects.push(p?); }

    if projects.is_empty() {
        println!("No projects found in logs to export.");
        return Ok(());
    }

    let selected_project = Select::new("Select Project to Export:", projects).prompt()?;

    // 2. Select Year and Month
    let now = Local::now();
    let default_year = now.year();
    let default_month = if now.month() == 1 { 12 } else { now.month() - 1 };

    let selected_year = CustomType::<i32>::new("Year:")
        .with_default(default_year)
        .prompt()?;
    
    let selected_month = CustomType::<u32>::new("Month (1-12):")
        .with_default(default_month)
        .prompt()?;

    let filename = format!("/Users/reneh/Downloads/Urenstaat_{}_{}_{}.xlsx", selected_year, selected_month, selected_project);

    let medewerker_name = env::var("EMPLOYEE_NAME").unwrap_or("John Doe".to_string());
    let medewerker_title = env::var("EMPLOYEE_TITLE").unwrap_or("Enterprise Architect".to_string());
    let medewerker_phone = env::var("EMPLOYEE_PHONE").unwrap_or("000000000".to_string());

    // 3. Create Workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    
    worksheet.protect();

    // Styles
    let title_fmt = Format::new().set_bold().set_font_size(14).set_align(FormatAlign::Left).set_font_name(FONT_NAME);

    let header_fmt = Format::new().set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10);
    let header_unlocked_fmt = Format::new().set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let header_address_fmt = Format::new().set_font_name(FONT_NAME).set_font_size(10);

    let sheet_header_fmt = Format::new().set_align(FormatAlign::Center).set_border(FormatBorder::Thin).set_background_color(Color::RGB(0xF28E00)).set_font_name(FONT_NAME).set_font_size(10);
    let sheet_description_fmt = Format::new().set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10);
    let sheet_hours_fmt = Format::new().set_align(FormatAlign::Center).set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10);
    let sheet_description_unlocked_fmt = Format::new().set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let sheet_hours_unlocked_fmt = Format::new().set_align(FormatAlign::Center).set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let sheet_total_description_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Left).set_font_name(FONT_NAME).set_font_size(10);
    let sheet_rowtotal_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Center).set_font_name(FONT_NAME).set_font_size(10);
    let sheet_daytotal_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Center).set_font_name(FONT_NAME).set_font_size(10);
    
    let header_expenses_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Left).set_font_name(FONT_NAME).set_font_size(10);
    let header_expenses_total_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Right).set_font_name(FONT_NAME).set_font_size(10);
    let expenses_date_fmt = Format::new().set_border(FormatBorder::Thin).set_num_format("dd-mm-yyyy").set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let expenses_description_fmt = Format::new().set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let expenses_amount_fmt = Format::new().set_num_format("€ #,##0.00").set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10);
    let expenses_amount_unlocked_fmt = Format::new().set_num_format("€ #,##0.00").set_border(FormatBorder::Thin).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let expenses_total_description_fmt = Format::new().set_font_name(FONT_NAME).set_font_size(10);
      
    let footer_header_fmt = Format::new().set_bold().set_align(FormatAlign::Left).set_font_name(FONT_NAME).set_font_size(10);
    let footer_fmt = Format::new().set_align(FormatAlign::Left).set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let footer_date_fmt = Format::new().set_num_format("dd-mm-yyyy").set_font_name(FONT_NAME).set_font_size(10).set_unlocked();
    let footer_signature_fmt = Format::new().set_bold().set_border(FormatBorder::Medium).set_align(FormatAlign::Top).set_font_name(FONT_NAME).set_font_size(10);
    
    // Layout
    worksheet.set_landscape();

    worksheet.set_print_area(0, 0, 45, 33)?;
    worksheet.set_print_gridlines(false);
    worksheet.set_paper_size(9); // A4
    worksheet.set_margins(0.25, 0.25, 0.5, 0.5, 0.25, 0.25);
    worksheet.set_print_fit_to_pages(1,1);

    for col in 0..=32 { worksheet.set_column_width(col, 6)?; }
    worksheet.set_column_width(1, 20)?; 
    worksheet.set_column_width(33, 10)?;

    // --- Write Headers ---
    worksheet.write_string_with_format(1, 1, "TIJDVERANTWOORDINGSFORMULIER", &title_fmt)?;

    worksheet.write_string_with_format(3, 1, "Naam medewerker", &header_fmt)?;
    worksheet.merge_range(3, 2, 3, 9, &medewerker_name,&header_fmt)?;

    worksheet.write_string_with_format(4, 1, "Functie in opdracht", &header_fmt)?;
    worksheet.merge_range(4, 2, 4, 9, &medewerker_title,&header_unlocked_fmt)?;
    worksheet.write_string_with_format(5, 1, "Telefoonnummer", &header_fmt)?;
    worksheet.merge_range(5, 2, 5, 9, &medewerker_phone,&header_unlocked_fmt)?;

    worksheet.write_string_with_format(7, 1, "Opdrachtgever", &header_fmt)?;
    worksheet.merge_range(7, 2, 7, 9, &selected_project,&header_unlocked_fmt)?;
    worksheet.write_string_with_format(8, 1, "Functie", &header_fmt)?;
    worksheet.merge_range(8, 2, 8, 9, "",&header_unlocked_fmt)?;
    worksheet.write_string_with_format(9, 1, "Projectnaam", &header_fmt)?;
    worksheet.merge_range(9, 2, 9, 9, "",&header_unlocked_fmt)?;
    worksheet.write_string_with_format(10, 1, "Projectnummer", &header_fmt)?;
    worksheet.merge_range(10, 2, 10, 9, "",&header_unlocked_fmt)?;
 
    let month_name_str = month_name(selected_month);
 
    worksheet.merge_range(3, 12, 3, 15, "Maand",&header_fmt)?;
    worksheet.merge_range(4, 12, 4, 15, "Jaar",&header_fmt)?;
    worksheet.merge_range(5, 12, 5, 15, "Invuldatum",&header_fmt)?;
    
    worksheet.merge_range(3, 16, 3, 20, month_name_str,&header_fmt)?;
    worksheet.merge_range(4, 16, 4, 20, &selected_year.to_string(),&header_fmt)?;
    worksheet.merge_range(5, 16, 5, 20, &now.format("%d-%m-%Y").to_string(),&header_unlocked_fmt)?;

    let logo_image = Image::new("logo.jpg")?;
    let logo_image = logo_image.set_scale_to_size(300, 200, true);

    worksheet.insert_image(2, 23, &logo_image)?;

    worksheet.write_string_with_format(7, 23, "Hoeksekade 160", &header_address_fmt)?;
    worksheet.write_string_with_format(8, 23, "2661, JL Bergschenhoek", &header_address_fmt)?;


    // --- Calendar Grid & Data Fetching ---
    let start_row_cal = 14; 
    let start_row_hours = 16;
    
    // Dutch short days
    let days_map = |w: Weekday| match w {
        Weekday::Mon => "Ma", Weekday::Tue => "Di", Weekday::Wed => "Wo",
        Weekday::Thu => "Do", Weekday::Fri => "Vr", Weekday::Sat => "Za", Weekday::Sun => "Zo"
    };

    // Prepare statement for fetching hours
    // We cannot parameterize column names, so we prepare the logic inside the loop
    
    worksheet.write_string_with_format(start_row_hours, 1, "Gewerkte uren", &sheet_description_fmt)?;

    for day in 1..=31 {
        let col_idx = day as u16;
        
        // Check if date is valid (e.g., Feb 30 will fail)
        if let Some(date) = NaiveDate::from_ymd_opt(selected_year, selected_month, day) {
            
            // Write Calendar Header
            worksheet.write_string_with_format(start_row_cal, col_idx+1, days_map(date.weekday()), &sheet_header_fmt)?;
            worksheet.write_number_with_format(start_row_cal + 1, col_idx+1, day, &sheet_header_fmt)?;

            // Fetch Data
            let iso_week = date.iso_week();
            // Create week string (Note: iso_week.year() handles year crossover, e.g., Dec 30 might be Week 1 of next year)
            let week_str = format!("{}-W{:02}", iso_week.year(), iso_week.week());
            
            // Map weekday to DB column
            let col_name = match date.weekday() {
                Weekday::Mon => "mon", Weekday::Tue => "tue", Weekday::Wed => "wed",
                Weekday::Thu => "thu", Weekday::Fri => "fri", Weekday::Sat => "sat", Weekday::Sun => "sun",
            };

            let sql = format!("SELECT {} FROM timesheets WHERE week = ?1 AND project = ?2", col_name);
            
            let hours: Option<f64> = conn.query_row(
                &sql, 
                params![week_str, selected_project], 
                |row| row.get(0)
            ).optional()?; // Returns Ok(None) if no row found

            let val = hours.unwrap_or(0.0);
            
            // Write Hours (Row 16, typically index 0 in the 5 blank rows)
            if val > 0.0 {
                worksheet.write_number_with_format(start_row_hours, col_idx+1, val, &sheet_hours_fmt)?;
            } else {
                worksheet.write_blank(start_row_hours, col_idx+1, &sheet_hours_fmt)?;
            }

        } else {
            // Invalid date (e.g. Feb 30), just format blank
             worksheet.write_blank(start_row_cal, col_idx+1, &sheet_header_fmt)?;
             worksheet.write_blank(start_row_cal+1, col_idx+1, &sheet_header_fmt)?;
             worksheet.write_blank(start_row_hours, col_idx+1, &sheet_hours_fmt)?;
        }
    }

    // Fill the remaining 4 empty rows just for visuals
    for r in 1..5 {
        worksheet.write_blank(start_row_hours + r, 1, &sheet_description_unlocked_fmt)?;
        for c in 2..=32 {
             worksheet.write_blank(start_row_hours + r, c, &sheet_hours_unlocked_fmt)?;
        }
    }

    // --- Totals Formulas ---
    worksheet.write_string_with_format(start_row_cal + 1, 33, "Totaal", &sheet_rowtotal_fmt)?;
    
    // Row Totals
    for r in 0..5 {
        let current_row = start_row_hours + r;
        let row_num_excel = current_row + 1;
        let formula = format!("=SUM(B{}:AF{})", row_num_excel, row_num_excel);
        worksheet.write_formula_with_format(current_row, 33, Formula::new(formula), &sheet_rowtotal_fmt)?;
    }

    // Column Totals
    let total_facturabel_row = start_row_hours + 5;
    worksheet.write_string_with_format(total_facturabel_row, 1, "Totaal facturabel", &sheet_total_description_fmt)?;

    for c in 2..=32 {
        let col_char = column_number_to_name(c); 
        let formula = format!("=SUM({}{}:{}{})", col_char, start_row_hours + 1, col_char, total_facturabel_row);
        worksheet.write_formula_with_format(total_facturabel_row, c, Formula::new(formula), &sheet_daytotal_fmt)?;
    }
    
    // Grand Total
    let formula_grand = format!("=SUM(AH{}:AH{})", start_row_hours + 1, total_facturabel_row);
    worksheet.write_formula_with_format(total_facturabel_row, 33, Formula::new(formula_grand), &sheet_rowtotal_fmt)?;

    // --- Expenses (Same as original) ---
    let expense_start_row = total_facturabel_row + 3;
    worksheet.write_string_with_format(expense_start_row, 1, "Onkostendeclaratie medewerker (bonnen bijvoegen)", &footer_header_fmt)?;
    let exp_header_row = expense_start_row + 1;

    worksheet.merge_range(exp_header_row, 1, exp_header_row, 2, "Datum", &header_expenses_fmt)?;
    worksheet.merge_range(exp_header_row, 3, exp_header_row, 22, "Omschrijving", &header_expenses_fmt)?;
    worksheet.merge_range(exp_header_row, 23, exp_header_row, 26, "Bedrag excl. BTW", &header_expenses_total_fmt)?;
    worksheet.merge_range(exp_header_row, 27, exp_header_row, 29, "BTW", &header_expenses_total_fmt)?;
    worksheet.merge_range(exp_header_row, 30, exp_header_row, 33, "Bedrag incl.", &header_expenses_total_fmt)?;

    for i in 0..4 {
        let r = exp_header_row + 1 + i;

        worksheet.merge_range(r, 1, r, 2, "", &expenses_date_fmt)?;
        worksheet.merge_range(r, 3, r, 22, "", &expenses_description_fmt)?;
        worksheet.merge_range(r, 23, r, 26, "", &expenses_amount_fmt)?;
        worksheet.merge_range(r, 27, r, 29, "", &expenses_amount_fmt)?;
        worksheet.merge_range(r, 30, r, 33, "", &expenses_amount_unlocked_fmt)?;

        worksheet.write_number_with_format(r, 23, 0, &expenses_amount_fmt)?; 
        worksheet.write_number_with_format(r, 27, 0, &expenses_amount_fmt)?; 
        worksheet.write_number_with_format(r, 27, 0, &expenses_amount_unlocked_fmt)?; 

        let row_excel = r + 1;
        let formula_incl = format!("=AE{}/121*100", row_excel);
        worksheet.write_formula_with_format(r, 23, Formula::new(formula_incl), &expenses_amount_fmt)?;
        let formula_incl = format!("=AE{}/121*21", row_excel);
        worksheet.write_formula_with_format(r, 27, Formula::new(formula_incl), &expenses_amount_fmt)?;
    }

    let exp_total_row = exp_header_row + 5;
    worksheet.write_string_with_format(exp_total_row, 3, "Totaal", &expenses_total_description_fmt)?;
    let start_sum = exp_header_row + 2; 
    let end_sum = exp_total_row;  
    worksheet.merge_range(exp_total_row, 23, exp_total_row, 26, "", &expenses_amount_fmt)?;
    worksheet.merge_range(exp_total_row, 27, exp_total_row, 29, "", &expenses_amount_fmt)?;
    worksheet.merge_range(exp_total_row, 30, exp_total_row, 33, "", &expenses_amount_fmt)?;      
    worksheet.write_formula_with_format(exp_total_row, 23, Formula::new(format!("=SUM(X{}:X{})", start_sum, end_sum)), &expenses_amount_fmt)?;
    worksheet.write_formula_with_format(exp_total_row, 27, Formula::new(format!("=SUM(AB{}:AB{})", start_sum, end_sum)), &expenses_amount_fmt)?;
    worksheet.write_formula_with_format(exp_total_row, 30, Formula::new(format!("=SUM(AE{}:AE{})", start_sum, end_sum)), &expenses_amount_fmt)?;

    // --- Signatures ---
    let sign_row = exp_total_row + 3;
    worksheet.write_string_with_format(sign_row, 1, "Opdrachtgever:", &footer_header_fmt)?;
    worksheet.write_string_with_format(sign_row + 1, 1, &selected_project, &footer_fmt)?;
    worksheet.write_string_with_format(sign_row + 2, 1, "Datum:", &footer_header_fmt)?;
    worksheet.write_string_with_format(sign_row + 3, 1, now.format("%d-%m-%Y").to_string(), &footer_date_fmt)?;


    worksheet.write_string_with_format(sign_row, 23, "Medewerker:", &footer_header_fmt)?;
    worksheet.write_string_with_format(sign_row + 1, 23, &medewerker_name, &footer_fmt)?;
    worksheet.write_string_with_format(sign_row + 2, 23, "Datum:", &footer_header_fmt)?;
    worksheet.write_string_with_format(sign_row + 3, 23, now.format("%d-%m-%Y").to_string(), &footer_date_fmt)?;


    worksheet.write_string_with_format(sign_row + 4, 1, "Handtekening opdrachtgever:", &footer_header_fmt)?;
    worksheet.write_string_with_format(sign_row + 4, 23, "Handtekening medewerker:", &footer_header_fmt)?;
    worksheet.set_row_height(sign_row+5, 120)?;
    worksheet.merge_range(sign_row+5, 1, sign_row+5, 9, "", &footer_signature_fmt)?;
    worksheet.merge_range(sign_row+5, 23, sign_row+5, 32, "", &footer_signature_fmt)?;  

    let signature = Image::new("signature.png")?;
    let signature = signature.set_scale_to_size(300, 200, true);

    worksheet.insert_image(sign_row + 5, 23, &signature)?;

    workbook.save(&filename)?;

    println!("File successfully generated: {}", filename);
    Ok(())
}

fn month_name(month_num: u32) -> &'static str {
    match month_num {
        1 => "Januari",
        2 => "Februari",
        3 => "Maart",
        4 => "April",
        5 => "Mei",
        6 => "Juni",
        7 => "Juli",
        8 => "Augustus",
        9 => "September",
        10 => "Oktober",
        11 => "November",
        12 => "December",
        _ => "Onbekend",
    }
}
