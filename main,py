import os
import math
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


# =========================================
# 기본 설정
# =========================================
POSITION_COL = "직책"
LEAVE_STATUS_COL = "휴직"

# 글로싸인 템플릿의 실제 헤더 행 번호(엑셀 기준 17행)
TEMPLATE_HEADER_ROW = 16  # pandas 기준 0부터 시작


# =========================================
# 공통 유틸
# =========================================
def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_phone(value):
    s = clean(value)
    if not s:
        return ""
    allowed = []
    for ch in s:
        if ch.isdigit() or ch in ["+", "-"]:
            allowed.append(ch)
    return "".join(allowed)


def ceil_to_ones(value):
    """
    1의 자리 올림
    예: 12341 -> 12350
    """
    if pd.isna(value) or clean(value) == "":
        return ""
    try:
        num = float(value)
        return int(math.ceil(num / 10.0) * 10)
    except Exception:
        return value


def num_to_str(value):
    if pd.isna(value) or clean(value) == "":
        return ""
    try:
        num = float(value)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except Exception:
        return clean(value)


def get_col(row, col_name, default=""):
    if col_name in row.index:
        return row[col_name]
    return default


# =========================================
# 데이터 처리
# =========================================
def load_payroll_df(path):
    df = pd.read_excel(path)
    df.columns = [clean(c) for c in df.columns]
    return df


def split_payroll(df):
    if POSITION_COL not in df.columns:
        raise ValueError(f"급여 파일에 '{POSITION_COL}' 열이 없습니다.")

    chair_df = df[df[POSITION_COL].astype(str).str.strip() == "주임"].copy()

    instructor_df = df[df[POSITION_COL].astype(str).str.strip() != "주임"].copy()

    if LEAVE_STATUS_COL in instructor_df.columns:
        instructor_df = instructor_df[
            instructor_df[LEAVE_STATUS_COL].fillna("").astype(str).str.strip() == ""
        ].copy()

    return chair_df, instructor_df


def load_template_columns(template_path):
    raw = pd.read_excel(template_path, header=None)

    if len(raw) <= TEMPLATE_HEADER_ROW:
        raise ValueError(
            f"{os.path.basename(template_path)} 파일에서 {TEMPLATE_HEADER_ROW + 1}행 헤더를 찾을 수 없습니다."
        )

    header_row = raw.iloc[TEMPLATE_HEADER_ROW].tolist()

    columns = []
    used = {}

    for col in header_row:
        name = clean(col)
        if not name:
            name = "Unnamed"

        if name in used:
            used[name] += 1
            name = f"{name}__dup{used[name]}"
        else:
            used[name] = 1

        columns.append(name)

    return columns


def build_common_mapping(pay_row, semester_value, contract_start_date, contract_end_date,
                         hr_manager_name, hr_manager_email, hr_manager_phone):
    mapping = {}

    # 급여 엑셀 열명과 같은 건 자동 복사
    for col in pay_row.index:
        mapping[col] = get_col(pay_row, col, "")

    # 명시적 보정
    mapping["이름"] = clean(get_col(pay_row, "이름", get_col(pay_row, "성명", "")))
    mapping["강사명(A)"] = clean(get_col(pay_row, "성명", get_col(pay_row, "이름", "")))
    mapping["이메일"] = clean(get_col(pay_row, "이메일", ""))
    mapping["휴대폰번호"] = normalize_phone(get_col(pay_row, "휴대폰번호", ""))

    mapping["해당학기"] = semester_value
    mapping["계약시작날짜"] = contract_start_date
    mapping["계약종료날짜"] = contract_end_date
    mapping["인사담당자"] = hr_manager_name
    mapping["인사담당자 이메일"] = hr_manager_email
    mapping["인사담당자 휴대폰번호"] = hr_manager_phone

    # 숫자 관련
    f_val = get_col(pay_row, "F", "")
    h_val = get_col(pay_row, "H", "")
    y_val = get_col(pay_row, "Y", "")
    aj_val = get_col(pay_row, "AJ", "")
    ak_val = get_col(pay_row, "AK", "")
    al_val = get_col(pay_row, "AL", "")
    am_val = get_col(pay_row, "AM", "")

    # FH = F + H
    try:
        fh_val = ""
        if clean(f_val) != "" or clean(h_val) != "":
            f_num = float(f_val) if clean(f_val) != "" else 0
            h_num = float(h_val) if clean(h_val) != "" else 0
            total = f_num + h_num
            fh_val = int(total) if float(total).is_integer() else total
    except Exception:
        fh_val = ""

    mapping["시수+간주시간(FH)"] = num_to_str(fh_val)
    mapping["시수(F)"] = num_to_str(f_val)
    mapping["간주시간(H)"] = num_to_str(h_val)
    mapping["확정 기본급(수정 1의자리 올림)(Y)"] = num_to_str(ceil_to_ones(y_val))
    mapping["총 표준 근로시간(AJ)"] = num_to_str(aj_val)
    mapping["총 수업준비 간주시간(AK)"] = num_to_str(ak_val)
    mapping["총시수(AL)"] = num_to_str(al_val)
    mapping["보정시수(AM)"] = num_to_str(am_val)

    return mapping


def row_to_template_record(pay_row, template_columns, semester_value, contract_start_date,
                           contract_end_date, hr_manager_name, hr_manager_email, hr_manager_phone):
    mapping = build_common_mapping(
        pay_row,
        semester_value,
        contract_start_date,
        contract_end_date,
        hr_manager_name,
        hr_manager_email,
        hr_manager_phone,
    )

    record = {}
    for col in template_columns:
        base_col = col.split("__dup")[0]
        if base_col in mapping:
            record[col] = mapping[base_col]
        else:
            record[col] = ""

    # 중복 이름/이메일/휴대폰번호 처리
    name_dup_cols = [c for c in template_columns if c.startswith("이름")]
    email_dup_cols = [c for c in template_columns if c.startswith("이메일")]
    phone_dup_cols = [c for c in template_columns if c.startswith("휴대폰번호")]

    if len(name_dup_cols) >= 2:
        record[name_dup_cols[0]] = clean(get_col(pay_row, "성명", get_col(pay_row, "이름", "")))
        record[name_dup_cols[1]] = hr_manager_name

    if len(email_dup_cols) >= 2:
        record[email_dup_cols[0]] = clean(get_col(pay_row, "이메일", ""))
        record[email_dup_cols[1]] = hr_manager_email

    if len(phone_dup_cols) >= 2:
        record[phone_dup_cols[0]] = normalize_phone(get_col(pay_row, "휴대폰번호", ""))
        record[phone_dup_cols[1]] = hr_manager_phone

    return record


def build_output_df(pay_df, template_columns, semester_value, contract_start_date,
                    contract_end_date, hr_manager_name, hr_manager_email, hr_manager_phone):
    records = []
    for _, row in pay_df.iterrows():
        records.append(
            row_to_template_record(
                row,
                template_columns,
                semester_value,
                contract_start_date,
                contract_end_date,
                hr_manager_name,
                hr_manager_email,
                hr_manager_phone,
            )
        )
    return pd.DataFrame(records, columns=template_columns)


def save_result(df, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    out_df = df.copy()
    out_df.columns = [c.split("__dup")[0] for c in out_df.columns]
    out_df.to_excel(output_path, index=False)


# =========================================
# GUI
# =========================================
class GloSignApp:
    def __init__(self, root):
        self.root = root
        self.root.title("글로싸인 대량전송 엑셀 자동생성기")
        self.root.geometry("820x700")
        self.root.resizable(True, True)

        self.payroll_path = tk.StringVar()
        self.chair_template_path = tk.StringVar()
        self.instructor_template_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))

        self.semester_value = tk.StringVar(value="261 봄학기")
        self.contract_start_date = tk.StringVar()
        self.contract_end_date = tk.StringVar()
        self.hr_manager_name = tk.StringVar()
        self.hr_manager_email = tk.StringVar()
        self.hr_manager_phone = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        title = ttk.Label(main, text="글로싸인 대량전송 엑셀 자동생성기", font=("Malgun Gothic", 16, "bold"))
        title.pack(anchor="w", pady=(0, 12))

        info = (
            "사용 순서\n"
            "1. 급여 엑셀 선택\n"
            "2. 주임용 / 강사용 템플릿 선택\n"
            "3. 학기, 계약기간, 인사담당자 정보 입력\n"
            "4. 생성 버튼 클릭\n"
            "5. output 폴더에서 결과 확인"
        )
        ttk.Label(main, text=info, justify="left").pack(anchor="w", pady=(0, 16))

        self.add_file_row(main, "급여 엑셀", self.payroll_path, self.browse_payroll)
        self.add_file_row(main, "주임용 템플릿", self.chair_template_path, self.browse_chair_template)
        self.add_file_row(main, "강사용 템플릿", self.instructor_template_path, self.browse_instructor_template)
        self.add_folder_row(main, "저장 폴더", self.output_dir, self.browse_output_dir)

        sep1 = ttk.Separator(main, orient="horizontal")
        sep1.pack(fill="x", pady=14)

        form = ttk.Frame(main)
        form.pack(fill="x")

        self.add_entry_row(form, "학기", self.semester_value, 0)
        self.add_entry_row(form, "계약시작날짜", self.contract_start_date, 1)
        self.add_entry_row(form, "계약종료날짜", self.contract_end_date, 2)
        self.add_entry_row(form, "인사담당자 이름", self.hr_manager_name, 3)
        self.add_entry_row(form, "인사담당자 이메일", self.hr_manager_email, 4)
        self.add_entry_row(form, "인사담당자 휴대폰", self.hr_manager_phone, 5)

        sep2 = ttk.Separator(main, orient="horizontal")
        sep2.pack(fill="x", pady=14)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(0, 10))

        ttk.Button(btn_frame, text="생성", command=self.run).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="입력값 초기화", command=self.clear_inputs).pack(side="left")

        ttk.Label(main, text="실행 로그", font=("Malgun Gothic", 11, "bold")).pack(anchor="w", pady=(10, 6))

        self.log_text = tk.Text(main, height=18, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

    def add_file_row(self, parent, label_text, var, command):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)

        ttk.Label(frame, text=label_text, width=16).pack(side="left")
        ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(frame, text="선택", command=command, width=10).pack(side="left")

    def add_folder_row(self, parent, label_text, var, command):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)

        ttk.Label(frame, text=label_text, width=16).pack(side="left")
        ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(frame, text="선택", command=command, width=10).pack(side="left")

    def add_entry_row(self, parent, label_text, var, row):
        ttk.Label(parent, text=label_text, width=16).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=var, width=50).grid(row=row, column=1, sticky="ew", padx=6, pady=5)
        parent.grid_columnconfigure(1, weight=1)

    def browse_payroll(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            self.payroll_path.set(path)

    def browse_chair_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            self.chair_template_path.set(path)

    def browse_instructor_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            self.instructor_template_path.set(path)

    def browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def clear_inputs(self):
        self.contract_start_date.set("")
        self.contract_end_date.set("")
        self.hr_manager_name.set("")
        self.hr_manager_email.set("")
        self.hr_manager_phone.set("")
        self.log("입력값 초기화 완료")

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.root.update_idletasks()

    def validate_inputs(self):
        if not self.payroll_path.get():
            messagebox.showwarning("확인", "급여 엑셀 파일을 선택하세요.")
            return False
        if not self.chair_template_path.get():
            messagebox.showwarning("확인", "주임용 템플릿 파일을 선택하세요.")
            return False
        if not self.instructor_template_path.get():
            messagebox.showwarning("확인", "강사용 템플릿 파일을 선택하세요.")
            return False
        if not self.output_dir.get():
            messagebox.showwarning("확인", "저장 폴더를 선택하세요.")
            return False
        return True

    def run(self):
        if not self.validate_inputs():
            return

        try:
            self.log("=" * 60)
            self.log("작업 시작")
            self.log("급여 파일 읽는 중...")
            payroll_df = load_payroll_df(self.payroll_path.get())

            self.log("주임/강사 분기 중...")
            chair_df, instructor_df = split_payroll(payroll_df)

            self.log(f"주임 대상: {len(chair_df)}명")
            self.log(f"강사 대상: {len(instructor_df)}명")

            self.log("주임용 템플릿 헤더 읽는 중...")
            chair_template_columns = load_template_columns(self.chair_template_path.get())

            self.log("강사용 템플릿 헤더 읽는 중...")
            instructor_template_columns = load_template_columns(self.instructor_template_path.get())

            self.log("주임용 결과 생성 중...")
            chair_output_df = build_output_df(
                chair_df,
                chair_template_columns,
                self.semester_value.get().strip(),
                self.contract_start_date.get().strip(),
                self.contract_end_date.get().strip(),
                self.hr_manager_name.get().strip(),
                self.hr_manager_email.get().strip(),
                self.hr_manager_phone.get().strip(),
            )

            self.log("강사용 결과 생성 중...")
            instructor_output_df = build_output_df(
                instructor_df,
                instructor_template_columns,
                self.semester_value.get().strip(),
                self.contract_start_date.get().strip(),
                self.contract_end_date.get().strip(),
                self.hr_manager_name.get().strip(),
                self.hr_manager_email.get().strip(),
                self.hr_manager_phone.get().strip(),
            )

            output_dir = Path(self.output_dir.get())
            output_dir.mkdir(parents=True, exist_ok=True)

            chair_output_path = output_dir / "대량전송_근로조건에 관한 확인서(주임용)_자동완성.xlsx"
            instructor_output_path = output_dir / "대량전송_근로조건에 관한 확인서(강사용)_자동완성.xlsx"

            self.log("파일 저장 중...")
            save_result(chair_output_df, str(chair_output_path))
            save_result(instructor_output_df, str(instructor_output_path))

            self.log("완료")
            self.log(f"주임용 저장: {chair_output_path}")
            self.log(f"강사용 저장: {instructor_output_path}")

            messagebox.showinfo(
                "완료",
                "자동완성 파일 생성이 끝났습니다.\n\n"
                f"주임용: {chair_output_path}\n"
                f"강사용: {instructor_output_path}"
            )

        except Exception as e:
            self.log(f"오류 발생: {e}")
            messagebox.showerror("오류", str(e))


# =========================================
# 실행
# =========================================
def main():
    root = tk.Tk()
    app = GloSignApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()