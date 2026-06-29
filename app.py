from __future__ import annotations

from pathlib import Path

import streamlit as st

from excel_tool import (
    CONFIG_PATH,
    SourceWorkbook,
    evaluate_rules,
    load_config,
    parse_mapping_rules,
    parse_product_rules,
    save_config,
    write_results_to_target,
    write_results_to_uploaded_target,
)


st.set_page_config(page_title="Excel产品数据自动提取工具", layout="wide")


def rows_from_editor(value) -> list[dict]:
    try:
        return value.to_dict("records")
    except AttributeError:
        return list(value or [])


if "config" not in st.session_state:
    st.session_state.config = load_config()

st.title("Excel产品数据自动提取与写入工具")
st.caption("批量上传源表，按客户/零件号/项目号识别产品，把取数或求和结果写入目标Excel。")

with st.sidebar:
    st.header("使用步骤")
    st.markdown(
        "1. 上传源表和目标表\n"
        "2. 配置产品规则\n"
        "3. 配置映射规则\n"
        "4. 先预览，再执行写入"
    )
    if st.button("重新加载已保存规则"):
        st.session_state.config = load_config()
        st.rerun()

st.subheader("1. 上传区")
source_uploads = st.file_uploader("上传源Excel文件，可多选", type=["xlsx"], accept_multiple_files=True)
target_path_text = st.text_input("目标Excel本机路径（填写后会覆盖该文件，并自动备份）", placeholder=r"C:\Users\你的名字\Desktop\目标表.xlsx")
target_upload = st.file_uploader("或上传目标Excel文件（不覆盖原文件，处理后下载）", type=["xlsx"], accept_multiple_files=False)

st.divider()
st.subheader("2. 产品规则区")
st.write("同一产品可以在客户、零件号、项目号中填写一个或多个值；多个值用逗号、分号或换行分隔。任一字段命中即识别为该产品。")
product_rows = st.data_editor(
    st.session_state.config["products"],
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "product_name": "产品名称",
        "customer_value": "客户关键值",
        "part_no_value": "零件号关键值",
        "project_no_value": "项目号关键值",
    },
    key="products_editor",
)

st.subheader("3. 映射规则区")
st.write("列配置可填写表头名称、Excel列字母（如 A）或列号（如 1）。")
mapping_rows = st.data_editor(
    st.session_state.config["mappings"],
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "rule_name": "规则名",
        "product_name": "产品名称",
        "source_sheet": "源Sheet",
        "customer_column": "客户列",
        "part_no_column": "零件号列",
        "project_no_column": "项目号列",
        "operation": st.column_config.SelectboxColumn("处理方式", options=["copy_value", "sum_column"]),
        "source_value_column": "取数/求和列",
        "target_sheet": "目标Sheet",
        "target_cell": "目标单元格",
        "empty_policy": st.column_config.SelectboxColumn("空值策略", options=["skip"]),
        "header_row": "表头行号",
    },
    key="mappings_editor",
)

col_save, col_preview, col_run = st.columns([1, 1, 1])

with col_save:
    if st.button("保存当前规则", type="secondary"):
        save_config(rows_from_editor(product_rows), rows_from_editor(mapping_rows), CONFIG_PATH)
        st.session_state.config = load_config()
        st.success(f"规则已保存到 {CONFIG_PATH.resolve()}")

products = parse_product_rules(rows_from_editor(product_rows))
mappings = parse_mapping_rules(rows_from_editor(mapping_rows))
sources = [SourceWorkbook(file.name, file.getvalue()) for file in source_uploads or []]

def show_logs(logs):
    if not logs:
        st.info("没有日志。")
        return
    records = [log.as_dict() for log in logs]
    st.dataframe(records, use_container_width=True)
    csv = "\ufeff" + "\n".join(
        [
            ",".join(["level", "rule_name", "product_name", "source_file", "source_sheet", "target_sheet", "target_cell", "value", "message"]),
            *[
                ",".join(str(row.get(key, "")).replace(",", "，") for key in ["level", "rule_name", "product_name", "source_file", "source_sheet", "target_sheet", "target_cell", "value", "message"])
                for row in records
            ],
        ]
    )
    st.download_button("下载日志CSV", csv.encode("utf-8-sig"), "processing_log.csv", "text/csv")

def validate_inputs() -> bool:
    if not sources:
        st.error("请至少上传一个源Excel文件。")
        return False
    if not target_path_text.strip() and target_upload is None:
        st.error("请填写目标Excel本机路径，或上传目标Excel文件。")
        return False
    if not products:
        st.error("请至少配置一个产品规则。")
        return False
    if not mappings:
        st.error("请至少配置一条映射规则。")
        return False
    return True

with col_preview:
    if st.button("预览匹配结果"):
        if validate_inputs():
            results = evaluate_rules(products, mappings, sources)
            preview_rows = []
            all_logs = []
            for result in results:
                preview_rows.append(
                    {
                        "规则名": result.mapping.rule_name,
                        "产品": result.mapping.product_name,
                        "处理方式": result.mapping.operation,
                        "匹配行数": result.matched_count,
                        "将写入": result.should_write,
                        "结果值": result.value,
                        "目标": f"{result.mapping.target_sheet}!{result.mapping.target_cell}",
                    }
                )
                all_logs.extend(result.logs)
            st.subheader("预览结果")
            st.dataframe(preview_rows, use_container_width=True)
            show_logs(all_logs)

with col_run:
    if st.button("执行写入", type="primary"):
        if validate_inputs():
            results = evaluate_rules(products, mappings, sources)
            if target_path_text.strip():
                target_path = Path(target_path_text).expanduser()
                try:
                    _, logs, backup_path = write_results_to_target(target_path, results, overwrite=True)
                    st.success(f"已覆盖目标文件，并生成备份：{backup_path}")
                    show_logs(logs)
                except Exception as exc:
                    st.error(f"写入失败：{exc}")
            else:
                try:
                    output_bytes, logs = write_results_to_uploaded_target(target_upload.getvalue(), results)
                    st.success("已生成处理后的目标文件。")
                    st.download_button(
                        "下载处理后的Excel",
                        output_bytes,
                        "updated_target.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    show_logs(logs)
                except Exception as exc:
                    st.error(f"写入失败：{exc}")
