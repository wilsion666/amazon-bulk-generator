from __future__ import annotations

import streamlit as st

from bulk_generator import run_generation


TEMPLATES = {
    "+ 新建关键词广告": """操作：新建广告
SKU：
广告活动：
广告组：
匹配方式：
竞价方式：固定竞价
预算：
竞价：
关键词：""",
    "+ 新建ASIN定投": """操作：新建ASIN定投
SKU：
广告活动：
广告组：
竞价方式：固定竞价
预算：
竞价：
ASIN：""",
    "+ 修改关键词竞价": """操作：修改关键词竞价
广告活动：
广告组：无
竞价：
关键词：""",
    "+ 修改ASIN竞价": """操作：修改ASIN定投竞价
广告活动：
广告组：无
竞价：
ASIN：""",
    "+ 修改预算": """操作：修改预算
广告活动：
预算：""",
    "+ 添加否定精准": """操作：添加否定精准
广告活动：
广告组：无
关键词：""",
    "+ 添加否定词组": """操作：添加否定词组
广告活动：
广告组：无
关键词：""",
    "+ 修改广告位": """操作：修改广告位百分比
广告活动：
首页首位：
商品页面：
其余位置：""",
    "+ 暂停广告活动": """操作：暂停广告活动
广告活动：""",
    "+ 暂停投放": """操作：暂停投放
广告活动：
广告组：无
关键词：
ASIN：""",
}

EXAMPLE_REQUEST = TEMPLATES["+ 新建关键词广告"]


st.set_page_config(
    page_title="亚马逊广告 Bulk 操作生成器",
    layout="wide",
)


def main() -> None:
    st.title("亚马逊广告 Bulk 操作生成器")

    bulk_template = _render_upload()
    submitted, requirement_text = _render_requirement()

    if submitted:
        result = run_generation(
            bulk_template=bulk_template,
            requirement_text=requirement_text,
        )
        _render_outputs(result)


def _render_upload():
    st.subheader("① 上传批量操作表格")
    return st.file_uploader(
        "Sponsored Products Bulk xlsx",
        type=["xlsx", "xlsm"],
        key="bulk_template",
    )


def _render_requirement() -> tuple[bool, str]:
    st.subheader("② 输入需求")
    if "requirement_text" not in st.session_state:
        st.session_state["requirement_text"] = EXAMPLE_REQUEST

    _render_template_buttons()
    _inject_editor_style_and_shortcuts()

    with st.form("bulk_request_form", clear_on_submit=False):
        requirement_text = st.text_area(
            "广告操作指令",
            key="requirement_text",
            height=320,
            label_visibility="collapsed",
        )
        st.caption("多个操作请用 --- 分隔；点击上方按钮可插入模板；Ctrl+Enter 解析预检查。")
        submitted = st.form_submit_button("生成 bulk_upload.xlsx", type="primary", width="stretch")
    return submitted, requirement_text


def _render_template_buttons() -> None:
    labels = list(TEMPLATES)
    for start in range(0, len(labels), 5):
        cols = st.columns(5)
        for col, label in zip(cols, labels[start : start + 5]):
            if col.button(label, key=f"tpl_{label}"):
                _append_template(TEMPLATES[label])


def _append_template(template: str) -> None:
    current = st.session_state.get("requirement_text", "").strip()
    if current:
        st.session_state["requirement_text"] = f"{current}\n\n---\n\n{template}"
    else:
        st.session_state["requirement_text"] = template


def _inject_editor_style_and_shortcuts() -> None:
    st.markdown(
        """
        <style>
        textarea {
            font-family: Consolas, "Courier New", monospace !important;
            line-height: 1.45 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.html(
        """
        <script>
        const install = () => {
          const doc = window.parent.document;
          const areas = doc.querySelectorAll('textarea');
          const area = areas[areas.length - 1];
          if (!area || area.dataset.bulkEditorReady === '1') return;
          area.dataset.bulkEditorReady = '1';
          area.addEventListener('keydown', (event) => {
            if (event.key === 'Tab') {
              event.preventDefault();
              const start = area.selectionStart;
              const end = area.selectionEnd;
              area.value = area.value.substring(0, start) + '\\t' + area.value.substring(end);
              area.selectionStart = area.selectionEnd = start + 1;
              area.dispatchEvent(new Event('input', { bubbles: true }));
            }
            if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
              event.preventDefault();
              const buttons = Array.from(doc.querySelectorAll('button'));
              const submit = buttons.find((button) => button.innerText.includes('生成 bulk_upload.xlsx'));
              if (submit) submit.click();
            }
          });
        };
        setTimeout(install, 250);
        </script>
        """,
        unsafe_allow_javascript=True,
    )


def _render_outputs(result) -> None:
    st.subheader("③ 生成 bulk_upload.xlsx")
    _render_generation_totals(result)

    if result.summary["bulk_ready"]:
        st.success("✅ bulk_upload.xlsx 已生成")
        _render_prechecks(result)
        st.download_button(
            "下载 bulk_upload.xlsx",
            data=result.bulk_upload,
            file_name="bulk_upload.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        return

    st.error("❌ bulk_upload.xlsx 未生成")
    st.write("失败原因：")
    for reason in _failure_reasons(result):
        st.markdown(f"- {reason}")
    _render_prechecks(result)


def _render_generation_totals(result) -> None:
    cols = st.columns(5)
    cols[0].metric("总操作块", result.summary.get("total_blocks", 0))
    cols[1].metric("成功操作块", result.summary.get("success_blocks", 0))
    cols[2].metric("部分成功", result.summary.get("partial_blocks", 0))
    cols[3].metric("失败操作块", result.summary.get("failed_blocks", 0))
    cols[4].metric("Bulk 行数", result.summary.get("generated_rows", 0))


def _render_prechecks(result) -> None:
    prechecks = result.summary.get("prechecks", [])
    if not prechecks:
        return
    st.write("预检查结果：")
    for item in prechecks:
        title = f"操作{item.get('block_index', '')}：{item.get('operation_type', '')}"
        status = item.get("status", "失败")
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.markdown(f"状态：{status}")
            st.markdown(f"广告活动：{item.get('campaign_name', '') or '未填写'}")
            st.markdown(f"广告组：{item.get('ad_group_name', '') or '未填写'}")
            if item.get("recognized_fields"):
                st.markdown(f"识别字段：{item.get('recognized_fields')}")
            if item.get("budget"):
                st.markdown(f"预算：{item.get('budget')}")
            if item.get("bid"):
                st.markdown(f"竞价：{item.get('bid')}")
            if item.get("keywords"):
                st.markdown(f"关键词：{item.get('keywords')}")
            if item.get("asins"):
                st.markdown(f"ASIN：{item.get('asins')}")
                st.markdown(f"成功匹配 ASIN：{item.get('matched_asins', '') or '无'}")
                st.markdown(f"未匹配 ASIN：{item.get('unmatched_asins', '') or '无'}")
            if item.get("placement"):
                st.markdown(f"广告位：{item.get('placement')} {item.get('percentage', '')}")
            st.markdown(f"成功匹配对象：{item.get('matched_objects', '') or '无'}")
            st.markdown(f"未匹配对象：{item.get('unmatched_objects', '') or '无'}")
            st.markdown(f"缺失字段：{item.get('missing_fields', '') or '无'}")
            if item.get("suggestion"):
                st.markdown(f"建议：{item.get('suggestion')}")
            st.markdown(f"预计生成行数：{item.get('generated_rows', 0)}")


def _failure_reasons(result) -> list[str]:
    prefix = "未生成 bulk_upload.xlsx："
    reasons: list[str] = []
    for raw_reason in result.summary.get("failure_reasons", []):
        reason = str(raw_reason).strip()
        if reason.startswith(prefix):
            reason = reason.removeprefix(prefix).strip()
        for part in reason.replace("；", ";").split(";"):
            clean_part = part.strip()
            if clean_part and clean_part not in reasons:
                reasons.append(clean_part)
    return reasons or ["未返回具体原因，请检查上传文件和需求文本。"]


if __name__ == "__main__":
    main()
