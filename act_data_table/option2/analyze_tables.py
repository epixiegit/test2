# -*- coding: utf-8 -*-
"""CSV/XLSX 컬럼 분석 + 실행 시 선택하는 컬럼별 차트.

1: Plotly HTML (반응형, 확대/이동/마우스오버, 공통 Plotly.js 참조)
2: Plotly PNG (plotly + kaleido + Chrome/Chromium/Edge 필요)
3: Matplotlib PNG (창 없이 이미지 저장)

입력: 이 파일 위치/targets (하위 폴더 포함)
출력: targets_analysis/원본 상대경로/컬럼분석.xlsx 및 컬럼차트/...
오류: error/YYYY-MM-DD_HH-MM.json에 JSON 배열로 누적 기록
한 번 실행에서 오류 기록 파일은 최대 5개. 같은 분의 오류는 같은 파일에 추가.
요약: targets_analysis/추가분석_실행요약.xlsx (이번 실행 기준으로 갱신)
실행수: 원본별 컬럼분석 XLSX 1건 + 컬럼 차트 각각 1건. 미지원 원본은 0건.
차트 비율의 분모는 결측치를 포함한 전체 원본 행 수.
숫자 차트의 IQR은 Q1~Q3이며 이상치 경계선은 아닙니다.
숫자 차트에는 Q2(중앙값)와 AVG(평균) 기준선도 표시합니다.
"""
from __future__ import annotations

import html
import json
import math
import os
import re
import shutil
import sys
import tempfile
import textwrap
import traceback
from collections import Counter
from datetime import datetime
from pathlib import Path

# 기존 Python 환경을 우선 사용하고, 프로젝트에 설치한 패키지도 보조 검색.
_LOCAL_PACKAGES = Path(__file__).resolve().parent / ".analysis_packages"
if _LOCAL_PACKAGES.is_dir():
    sys.path.append(str(_LOCAL_PACKAGES))

# 선택하지 않은 차트 라이브러리 때문에 프로그램이 중단되지 않도록 지연 로딩.
_COMMON_IMPORT_ERROR = None
try:
    import pandas as pd
    import numpy as np
    import openpyxl
    from pandas.api.types import is_bool_dtype, is_complex_dtype, is_numeric_dtype
except ImportError as exc:
    _COMMON_IMPORT_ERROR = exc

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "targets"
OUTPUT_DIR = BASE_DIR / "targets_analysis"
ERROR_DIR = BASE_DIR / "error"
MAX_ERROR_FILES_PER_RUN = 5
_ERROR_FILES_THIS_RUN: set[Path] = set()
_ERROR_LIMIT_ANNOUNCED = False
PLOTLY_JS_PATH = OUTPUT_DIR / "_assets" / "plotly.min.js"
_PLOTLY_JS_READY = False
TABLE_EXTENSIONS = {".csv", ".xlsx"}
RESULT_COLUMNS = [
    "컬럼명", "데이터타입", "고유값수", "결측치수", "결측률(%)",
    "최대값", "최소값", "평균값", "중앙값", "고유값목록",
]
CHART_OPTIONS = {
    "1": {"label": "Plotly HTML", "extension": ".html"},
    "2": {"label": "Plotly PNG", "extension": ".png"},
    "3": {"label": "Matplotlib PNG", "extension": ".png"},
}


def reset_error_file_limit() -> None:
    """폴더의 과거 파일 수와 무관하게 이번 실행의 한도를 초기화합니다."""
    global _ERROR_LIMIT_ANNOUNCED
    _ERROR_FILES_THIS_RUN.clear()
    _ERROR_LIMIT_ANNOUNCED = False


def can_write_error_file(path: Path) -> bool:
    """한도 내 새 파일 또는 이번 실행에서 이미 사용한 파일만 허용."""
    global _ERROR_LIMIT_ANNOUNCED
    if path.resolve() in _ERROR_FILES_THIS_RUN:
        return True
    if len(_ERROR_FILES_THIS_RUN) < MAX_ERROR_FILES_PER_RUN:
        return True
    if not _ERROR_LIMIT_ANNOUNCED:
        print(
            f"[오류 파일 한도] 이번 실행의 {MAX_ERROR_FILES_PER_RUN}개 한도에 도달했습니다. "
            "추가 오류의 새 파일 저장은 생략하고 콘솔에만 표시합니다."
        )
        _ERROR_LIMIT_ANNOUNCED = True
    return False


def log_error(exc: Exception, parameters: dict, stage: str = "chart"):
    """분 단위 JSON 배열에 누적. 한 실행에서 최대 5개 파일까지만 기록."""
    now = datetime.now().astimezone()
    record = {
        "date": now.date().isoformat(),
        "timestamp": now.isoformat(timespec="milliseconds"),
        "stage": stage,
        "parameters": parameters,
        "error": {
            "type": type(exc).__name__,
            "message": str(exc),
            "traceback": "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
        },
    }
    temporary = None
    try:
        path = ERROR_DIR / f"{now:%Y-%m-%d_%H-%M}.json"
        recovered_path = ERROR_DIR / f"{now:%Y-%m-%d_%H-%M}_recovered.json"
        if recovered_path.resolve() in _ERROR_FILES_THIS_RUN:
            path = recovered_path
        if not can_write_error_file(path):
            return None
        ERROR_DIR.mkdir(parents=True, exist_ok=True)
        entries = []
        if path.exists():
            try:
                entries = json.loads(path.read_text(encoding="utf-8-sig"))
                if not isinstance(entries, list):
                    raise ValueError("기존 오류 JSON이 배열이 아닙니다.")
            except (ValueError, UnicodeError):
                # 손상된 기록을 덮어쓰지 않습니다. 새 파일의 배열에 이어 기록.
                path = ERROR_DIR / f"{now:%Y-%m-%d_%H-%M}_recovered.json"
                if not can_write_error_file(path):
                    return None
                if path.exists():
                    entries = json.loads(path.read_text(encoding="utf-8-sig"))
                    if not isinstance(entries, list):
                        raise ValueError("복구 오류 JSON이 배열이 아닙니다.")
                else:
                    entries = []
        entries.append(record)
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=ERROR_DIR,
            prefix="_error_", suffix=".tmp", delete=False,
        ) as stream:
            temporary = Path(stream.name)
            json.dump(entries, stream, ensure_ascii=False, indent=2, default=str)
            stream.write("\n")
        temporary.replace(path)
        _ERROR_FILES_THIS_RUN.add(path.resolve())
        print(f"[오류 기록] {path}")
        return path
    except Exception as logging_error:
        # 기록 자체가 불가능해도 실제 오류를 콘솔에서 확인할 수 있게 합니다.
        print(f"[오류 기록 실패] {logging_error}")
        print(json.dumps(record, ensure_ascii=False, default=str))
        return None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def choose_chart_option():
    print("\n차트 출력 옵션")
    for option, info in CHART_OPTIONS.items():
        print(f"  {option}. {info['label']}")
    while True:
        try:
            option = input("번호 입력 [1/2/3]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n선택을 취소했습니다.")
            return None
        if option in CHART_OPTIONS:
            return option
        print("1, 2, 3 중 하나를 입력해 주세요.")


def configure_chart_backend(option: str) -> None:
    """선택한 엔진만 import. PNG용 브라우저는 설치된 실행 파일만 사용."""
    if option in {"1", "2"}:
        global go
        import plotly.graph_objects as go
        if option == "1":
            prepare_plotly_library()
        if option == "2":
            import kaleido  # PNG 의존성 확인
            if not os.environ.get("BROWSER_PATH"):
                candidates = [
                    shutil.which("chrome"), shutil.which("chromium"),
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                    str(Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe"),
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                ]
                browser = next((str(p) for p in candidates if p and Path(p).is_file()), None)
                if not browser:
                    raise RuntimeError(
                        "Plotly PNG 저장에 Chrome/Chromium/Edge가 필요합니다. "
                        "설치된 브라우저를 확인하거나 BROWSER_PATH를 지정해 주세요."
                    )
                os.environ["BROWSER_PATH"] = browser
    else:
        global plt, ticker
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager, ticker
        available = {font.name for font in font_manager.fontManager.ttflist}
        for name in ("Malgun Gothic", "NanumGothic", "AppleGothic", "DejaVu Sans"):
            if name in available:
                plt.rcParams["font.family"] = name
                break
        plt.rcParams["axes.unicode_minus"] = False
        plt.rcParams["axes.formatter.useoffset"] = False
        plt.rcParams["text.parse_math"] = False


def safe_filename(value) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value))
    return text.strip(" .")[:45] or "이름없음"


def numeric_column(series: pd.Series) -> bool:
    return (
        is_numeric_dtype(series.dtype)
        and not is_bool_dtype(series.dtype)
        and not is_complex_dtype(series.dtype)
    )


def chart_label(value) -> str:
    text = repr(value) if isinstance(value, str) else str(value)
    lines = textwrap.wrap(text, width=22, replace_whitespace=False) or [text]
    return "\n".join(lines[:3]) + ("…" if len(lines) > 3 else "")


def prepare_chart_data(series: pd.Series, sheet_name: str, column_name) -> dict:
    """3개 출력 방식이 공유하는 빈도·비율·히스토그램·IQR 계산."""
    total = len(series)
    missing = int(series.isna().sum())
    model = {
        "title": f"{sheet_name} · {column_name}",
        "total": total, "missing": missing, "excluded": 0,
        "counts": [], "shares": [], "labels": [], "full_labels": [],
        "q1": None, "q2": None, "q3": None, "avg": None, "iqr": None,
        "kind": "numeric" if numeric_column(series) else "categorical",
    }
    if model["kind"] == "categorical":
        counts = series.dropna().value_counts(sort=True)
        counts = counts[counts > 0]
        top = counts.head(20)
        model["counts"] = [int(n) for n in top.to_numpy()]
        model["shares"] = [n / total * 100 if total else 0 for n in model["counts"]]
        model["labels"] = [chart_label(value) for value in top.index]
        model["full_labels"] = [repr(v) if isinstance(v, str) else str(v) for v in top.index]
        model["width"] = int(max(1200, min(2400, len(top) * 95)))
        shown = int(top.sum())
        model["description"] = (
            f"전체 {total:,}행 · 결측 {missing:,}행 · 고유값 {len(counts):,}개\n"
            f"상위 20개 표시 {shown:,}행 · 나머지 값 {int(counts.iloc[20:].sum()):,}행. "
            "비율 분모는 전체 행 수이며 결측값은 막대에서 제외합니다."
        )
    else:
        raw = series.dropna().to_numpy(dtype=float)
        finite = raw[np.isfinite(raw)]
        model["excluded"] = len(raw) - len(finite)
        model["width"] = 1600
        if len(finite):
            bins = 1 if np.all(finite == finite[0]) else min(30, max(5, int(math.sqrt(len(finite)))))
            counts, edges = np.histogram(finite, bins=bins)
            if not np.isfinite(edges).all():
                raise ValueError("숫자 범위가 너무 커서 유효한 히스토그램 구간을 만들 수 없습니다.")
            model["counts"] = [int(n) for n in counts]
            model["shares"] = [n / total * 100 for n in model["counts"]]
            model["centers"] = (edges[:-1] / 2 + edges[1:] / 2).tolist()
            model["widths"] = np.diff(edges).tolist()
            model["edges"] = edges.tolist()
            q1, q2, q3 = np.quantile(finite, [0.25, 0.5, 0.75])
            # 먼저 나눠 합산해 큰 유한값들의 합계가 넘치는 것을 줄입니다.
            average = float(np.sum(finite / len(finite)))
            model.update(q1=float(q1), q2=float(q2), q3=float(q3),
                         avg=average, iqr=float(q3 - q1))
            if not all(math.isfinite(model[key]) for key in ("q1", "q2", "q3", "avg", "iqr")):
                raise ValueError("숫자 범위가 너무 커서 사분위수 또는 평균을 계산할 수 없습니다.")
        model["description"] = (
            f"전체 {total:,}행 · 표시 {len(finite):,}행 · 결측 {missing:,}행 · "
            f"무한대 제외 {model['excluded']:,}행\n"
            "막대 위: 구간 개수와 전체 행 대비 비율. IQR은 Q1~Q3 구간입니다."
        )
    return model


def build_plotly_figure(model: dict):
    """Plotly HTML/PNG가 공통으로 사용하는 Figure."""
    fig = go.Figure()
    categorical = model["kind"] == "categorical"
    counts = model["counts"]
    if counts:
        positions = list(range(len(counts))) if categorical else model["centers"]
        labels = [
            (f"{count:,} / {model['total']:,}행<br>{share:.1f}%" if categorical
             else f"{count:,}<br>{share:.1f}%") if count else ""
            for count, share in zip(counts, model["shares"])
        ]
        kwargs = {"width": 0.72} if categorical else {"width": [w * 0.98 for w in model["widths"]]}
        fig.add_trace(go.Bar(
            x=positions, y=counts, marker_color="#248a77",
            text=labels, textposition="outside", cliponaxis=False,
            customdata=[
                [html.escape(model["full_labels"][i]) if categorical
                 else f"{model['edges'][i]:,.6g} ~ {model['edges'][i+1]:,.6g}", share]
                for i, share in enumerate(model["shares"])
            ],
            hovertemplate="%{customdata[0]}<br>개수: %{y:,}<br>전체 대비: %{customdata[1]:.1f}%<extra></extra>",
            **kwargs,
        ))
        if categorical:
            fig.update_xaxes(
                tickmode="array", tickvals=positions,
                ticktext=[html.escape(label).replace("\n", "<br>") for label in model["labels"]],
                tickangle=-35,
            )
        else:
            q1, q3 = model["q1"], model["q3"]
            fig.add_vrect(x0=q1, x1=q3, fillcolor="#e8b04c", opacity=0.18, line_width=0, layer="below")
            if q1 == q3:
                fig.add_vline(x=q1, line_dash="dash", line_color="#c87918",
                              annotation_text=f"Q1=Q3={q1:,.6g}", annotation_position="top left")
            else:
                fig.add_vline(x=q1, line_dash="dash", line_color="#c87918",
                              annotation_text=f"Q1={q1:,.6g}", annotation_position="top left")
                fig.add_vline(x=q3, line_dash="dash", line_color="#8759a7",
                              annotation_text=f"Q3={q3:,.6g}", annotation_position="top right")
            q2, average = model["q2"], model["avg"]
            if q2 == average:
                fig.add_vline(
                    x=q2, line_color="#2563eb", line_width=2.2,
                    annotation_text=f"Q2(중앙값)=AVG(평균)={q2:,.6g}",
                    annotation_position="bottom left", annotation_yshift=10,
                )
            else:
                fig.add_vline(
                    x=q2, line_color="#2563eb", line_width=2.2,
                    annotation_text=f"Q2(중앙값)={q2:,.6g}",
                    annotation_position="bottom left", annotation_yshift=10,
                )
                fig.add_vline(
                    x=average, line_color="#dc2626", line_dash="dot", line_width=2.2,
                    annotation_text=f"AVG(평균)={average:,.6g}",
                    annotation_position="bottom right", annotation_yshift=10,
                )
            fig.add_annotation(
                x=1, y=1.16, xref="paper", yref="paper", xanchor="right",
                text=f"IQR = {model['iqr']:,.6g}", showarrow=False,
            )
        fig.update_yaxes(range=[0, max(1, max(counts)) * 1.35], dtick=1 if max(counts) < 10 else None)
    else:
        fig.add_annotation(
            x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False,
            text="표시할 비결측 값이 없습니다." if categorical else "표시할 유한한 숫자 값이 없습니다.",
        )
        fig.update_xaxes(showticklabels=False)
        fig.update_yaxes(showticklabels=False)
    title = html.escape(textwrap.fill(model["title"], width=90)).replace("\n", "<br>")
    subtitle = "값별 빈도 · 상위 20개" if categorical else "숫자 분포"
    fig.update_layout(
        template="plotly_white", autosize=True, showlegend=False,
        title={"text": title + "<br><sup>" + subtitle + "</sup>", "x": 0.02},
        font={"family": "Malgun Gothic, AppleGothic, sans-serif", "size": 12},
        xaxis_title="고유값 (원문 표기 기준)" if categorical else "값 (구간)",
        yaxis_title="등장 횟수 (행)",
        margin={"l": 85, "r": 60, "t": 140, "b": 245},
        paper_bgcolor="white", plot_bgcolor="white",
        hovermode="closest",
    )
    fig.add_annotation(
        x=0, y=-0.42, xref="paper", yref="paper", xanchor="left", yanchor="top",
        text=html.escape(model["description"]).replace("\n", "<br>"),
        showarrow=False, align="left", font={"size": 11},
    )
    return fig


def prepare_plotly_library() -> Path:
    """큰 JS 라이브러리는 출력 루트에 한 번만 저장합니다."""
    global _PLOTLY_JS_READY
    if _PLOTLY_JS_READY and PLOTLY_JS_PATH.is_file():
        return PLOTLY_JS_PATH
    from plotly.offline import get_plotlyjs
    library = get_plotlyjs()
    PLOTLY_JS_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        if not PLOTLY_JS_PATH.is_file() or PLOTLY_JS_PATH.read_text(encoding="utf-8") != library:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=PLOTLY_JS_PATH.parent,
                prefix="_plotly_", suffix=".js", delete=False,
            ) as stream:
                temporary = Path(stream.name)
                stream.write(library)
            temporary.replace(PLOTLY_JS_PATH)
        _PLOTLY_JS_READY = True
        return PLOTLY_JS_PATH
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def create_plotly_html_chart(model: dict, path: Path) -> None:
    library_path = prepare_plotly_library()
    fig = build_plotly_figure(model)
    document = fig.to_html(
        full_html=True, include_plotlyjs=False, auto_play=False,
        default_width="100%", default_height="100%",
        config={"responsive": True, "displaylogo": False, "scrollZoom": True},
    )
    style = "<style>html,body{width:100%;height:100%;margin:0;min-height:650px}</style>"
    relative_js = os.path.relpath(library_path, start=path.parent).replace(os.sep, "/")
    shared_script = '<script src="' + html.escape(relative_js, quote=True) + '"></script>'
    document = document.replace("<head>", '<head><meta charset="utf-8">' + style + shared_script, 1)
    path.write_text(document, encoding="utf-8")


def create_plotly_png_chart(model: dict, path: Path) -> None:
    fig = build_plotly_figure(model)
    fig.write_image(str(path), format="png", width=model["width"], height=900, scale=1.25)


def create_matplotlib_png_chart(model: dict, path: Path) -> None:
    fig = None
    try:
        fig, ax = plt.subplots(figsize=(model["width"] / 100, 8.5))
        categorical = model["kind"] == "categorical"
        counts = model["counts"]
        if not counts:
            ax.text(0.5, 0.5, "표시할 유효 값이 없습니다.",
                    ha="center", va="center", transform=ax.transAxes)
            ax.set_xticks([])
            ax.set_yticks([])
        else:
            positions = list(range(len(counts))) if categorical else model["centers"]
            widths = 0.72 if categorical else model["widths"]
            bars = ax.bar(positions, counts, width=widths, color="#248a77", edgecolor="white")
            if categorical:
                ax.set_xticks(positions)
                ax.set_xticklabels(model["labels"], rotation=40, ha="right", fontsize=9)
            else:
                q1, q3 = model["q1"], model["q3"]
                ax.axvspan(q1, q3, color="#e8b04c", alpha=0.18, label=f"IQR = {model['iqr']:,.6g}")
                ax.axvline(q1, color="#c87918", linestyle="--", label=f"Q1 = {q1:,.6g}")
                ax.axvline(q3, color="#8759a7", linestyle="--", label=f"Q3 = {q3:,.6g}")
                q2, average = model["q2"], model["avg"]
                if q2 == average:
                    ax.axvline(q2, color="#2563eb", linewidth=2.2,
                               label=f"Q2(중앙값)=AVG(평균) = {q2:,.6g}")
                else:
                    ax.axvline(q2, color="#2563eb", linewidth=2.2,
                               label=f"Q2(중앙값) = {q2:,.6g}")
                    ax.axvline(average, color="#dc2626", linestyle=":", linewidth=2.2,
                               label=f"AVG(평균) = {average:,.6g}")
                ax.legend(loc="upper right", fontsize=9)
                total = model["total"]
                right = ax.secondary_yaxis(
                    "right", functions=(lambda n: n / total * 100, lambda p: p * total / 100)
                )
                right.set_ylabel("전체 행 대비 비율")
                right.yaxis.set_major_formatter(ticker.PercentFormatter(xmax=100, decimals=1))
            for bar, count, share in zip(bars, counts, model["shares"]):
                if count:
                    count_text = f"{count:,} / {model['total']:,}행" if categorical else f"{count:,}"
                    ax.annotate(
                        f"{count_text}\n{share:.1f}%",
                        (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                        xytext=(0, 5), textcoords="offset points", ha="center", va="bottom", fontsize=8,
                    )
            ax.set_ylim(0, max(1, max(counts)) * 1.35)
            ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
        subtitle = "값별 빈도 · 상위 20개" if categorical else "숫자 분포"
        ax.set_title(textwrap.fill(model["title"], width=90) + "\n" + subtitle, fontsize=14, pad=18)
        ax.set_xlabel("고유값 (원문 표기 기준)" if categorical else "값 (구간)")
        ax.set_ylabel("등장 횟수 (행)")
        ax.grid(axis="y", alpha=0.16)
        ax.set_axisbelow(True)
        fig.subplots_adjust(bottom=0.3, top=0.84, right=0.88)
        fig.text(0.02, 0.02, model["description"], fontsize=9, va="bottom")
        fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    finally:
        if fig is not None:
            plt.close(fig)


def column_charts(table: pd.DataFrame, sheet_name: str, destination: Path,
                  option: str, source: Path, run_stats: dict | None = None):
    """공통 데이터 계산 후 선택한 출력 함수로 전달. 실패는 컬럼별 JSON 누적."""
    if run_stats is None:
        run_stats = new_run_stats()
    destination.mkdir(parents=True, exist_ok=True)
    renderers = {
        "1": create_plotly_html_chart,
        "2": create_plotly_png_chart,
        "3": create_matplotlib_png_chart,
    }
    generated, failed = [], 0
    for index, name in enumerate(table.columns, start=1):
        series = table.iloc[:, index - 1]
        path = destination / f"col_{index:03d}_{safe_filename(name)}{CHART_OPTIONS[option]['extension']}"
        parameters = {
            "option": option, "chart_engine": CHART_OPTIONS[option]["label"],
            "source_file": str(source), "sheet": sheet_name,
            "column": str(name), "column_index": index,
            "dtype": str(series.dtype), "rows": len(series),
            "output_file": str(path), "top_n": 20, "maximum_bins": 30,
            "percentage_denominator": "원본 전체 행 수 (결측 포함)",
            "iqr_lines": ["Q1", "Q3"],
            "reference_lines": ["Q1", "Q2(중앙값)", "Q3", "AVG(평균)"],
        }
        temporary = None
        run_stats["attempted"] += 1
        try:
            model = prepare_chart_data(series, sheet_name, name)
            parameters.update({
                "chart_type": model["kind"], "missing": model["missing"],
                "excluded_infinite": model["excluded"],
                "displayed_groups_or_bins": len(model["counts"]),
                "q1": model["q1"], "q2": model["q2"], "q3": model["q3"],
                "avg": model["avg"], "iqr": model["iqr"],
                "png_width": model["width"], "png_height": 900,
            })
            with tempfile.NamedTemporaryFile(
                dir=destination, prefix="_chart_", suffix=path.suffix, delete=False
            ) as stream:
                temporary = Path(stream.name)
            renderers[option](model, temporary)
            if not temporary.is_file() or temporary.stat().st_size == 0:
                raise RuntimeError("차트 파일이 생성되지 않았거나 파일이 비어 있습니다.")
            temporary.replace(path)
            generated.append(path)
            run_stats["completed"] += 1
            run_stats["chart_completed"] += 1
        except KeyboardInterrupt as exc:
            run_stats["errors"] += 1
            run_stats["chart_errors"] += 1
            log_error(exc, parameters, stage="chart_interrupted")
            raise
        except Exception as exc:
            failed += 1
            run_stats["errors"] += 1
            run_stats["chart_errors"] += 1
            print(f"[차트 실패] {sheet_name} / {name}: {type(exc).__name__}: {exc}")
            log_error(exc, parameters, stage="chart_generation")
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
    return generated, failed


def remove_old_charts(root: Path, generated: list[Path]) -> None:
    """전체 차트 성공 후에만 이전 선택 옵션/삭제된 컬럼의 생성물을 정리."""
    expected = {path.resolve() for path in generated}
    resolved_root = root.resolve()
    for path in root.glob("sheet_*/col_*"):
        if path.suffix.lower() not in {".png", ".html"}:
            continue
        if path.is_symlink() or not path.resolve().is_relative_to(resolved_root):
            continue
        if not re.fullmatch(r"sheet_\d+_.+", path.parent.name):
            continue
        if not re.fullmatch(r"col_\d+_.+\.(png|html)", path.name):
            continue
        if path.resolve() not in expected:
            path.unlink()


def read_csv(path: Path) -> pd.DataFrame:
    """UTF-8(BOM 포함), 한국어 Windows CP949, BOM 있는 UTF-16 지원."""
    with path.open("rb") as stream:
        prefix = stream.read(4)
    if prefix.startswith((b"\xff\xfe", b"\xfe\xff")):
        return pd.read_csv(path, encoding="utf-16")

    last_error = None
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(
        "CSV 인코딩을 읽지 못했습니다. UTF-8 BOM 또는 CP949로 저장해 주세요."
    ) from last_error


def safe_stat(series: pd.Series, operation: str):
    """계산 실패, 비어 있는 결과, NaN/무한대는 '-'로 대체."""
    try:
        value = getattr(series, operation)()
        if pd.isna(value) or not math.isfinite(value):
            return "-"
        return value
    except (TypeError, ValueError, OverflowError, ArithmeticError):
        return "-"


def unique_text(series: pd.Series, unique_count: int) -> str:
    if unique_count == 0 or unique_count > 10:
        return "-"
    # 최초 등장 순서 유지. pd.read_*의 기본 결측치 판정 규칙을 따릅니다.
    values = series.dropna().unique()
    text = ", ".join(str(value) for value in values)
    # Excel 셀 하나의 최대 텍스트 길이 제한.
    if len(text) > 32767:
        suffix = " ... (Excel 셀 길이 제한으로 생략)"
        text = text[:32767 - len(suffix)] + suffix
    return text


def analyze_columns(table: pd.DataFrame) -> pd.DataFrame:
    records = []
    for index, column_name in enumerate(table.columns):
        series = table.iloc[:, index]
        count = int(series.nunique(dropna=True))
        missing_count = int(series.isna().sum())
        missing_rate = (
            round(missing_count / len(series) * 100, 1)
            if len(series) else "-"
        )
        numeric = numeric_column(series)
        statistics = {
            label: safe_stat(series, method) if numeric else "-"
            for label, method in (
                ("최대값", "max"), ("최소값", "min"),
                ("평균값", "mean"), ("중앙값", "median"),
            )
        }
        records.append({
            "컬럼명": str(column_name),
            "데이터타입": str(series.dtype),
            "고유값수": count,
            "결측치수": missing_count,
            "결측률(%)": missing_rate,
            **statistics,
            "고유값목록": unique_text(series, count),
        })
    return pd.DataFrame(records, columns=RESULT_COLUMNS)


def source_tables(path: Path):
    if path.suffix.lower() == ".csv":
        yield "컬럼분석", read_csv(path)
    else:
        # 모든 시트를 차례로 읽어 한꺼번에 메모리에 올리지 않습니다.
        with pd.ExcelFile(path, engine="openpyxl") as book:
            for sheet_name in book.sheet_names:
                yield sheet_name, pd.read_excel(book, sheet_name=sheet_name)


def format_sheet(sheet) -> None:
    """헤더 고정, 필터, 너비 지정 및 원문 문자열의 수식 해석 방지."""
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    widths = {"A": 28, "B": 22, "C": 14, "D": 14, "E": 14,
              "F": 20, "G": 20, "H": 20, "I": 20, "J": 65}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"  # '=...'인 컬럼명/값도 수식 아닌 문자
    for cell in sheet[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill(
            fill_type="solid", fgColor="087D73"
        )
    for row in sheet.iter_rows(min_row=2, min_col=5, max_col=9):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                # 결측률은 0~100 숫자이며, 헤더에 % 단위를 표시합니다.
                cell.number_format = (
                    "0.0" if cell.column == 5 else "#,##0.######"
                )


def new_run_stats() -> dict:
    return {
        "attempted": 0, "completed": 0, "errors": 0,
        "chart_completed": 0, "chart_errors": 0,
    }


def save_execution_summary(stats_by_source: dict[Path, dict]) -> Path:
    """이번 실행의 원본별 집계를 하나의 XLSX로 저장(재실행 시 갱신)."""
    # 공통 라이브러리/차트 준비에 실패해도 openpyxl만 있으면 요약 저장 가능.
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    path = OUTPUT_DIR / "추가분석_실행요약.xlsx"
    headers = ["원본파일경로", "파일명", "추가분석실행수", "완료수", "에러수"]
    book = Workbook()
    sheet = book.active
    sheet.title = "실행요약"
    sheet.append(headers)
    for source, stats in stats_by_source.items():
        sheet.append([
            str(source), source.name,
            stats["attempted"], stats["completed"], stats["errors"],
        ])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 100, "B": 45, "C": 20, "D": 14, "E": 14}.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="087D73")
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=5):
        for cell in row:
            cell.number_format = "#,##0"

    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=OUTPUT_DIR, prefix="_execution_summary_", suffix=".xlsx", delete=False
        ) as stream:
            temporary = Path(stream.name)
        book.save(temporary)
        temporary.replace(path)
        return path
    finally:
        book.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def table_analysis(source: Path, chart_option: str,
                   run_stats: dict | None = None) -> tuple[Path, int, int, int]:
    """실행수 = XLSX 1건 + 시트별 컬럼 차트 건수. 완료는 파일 저장 성공 기준."""
    if run_stats is None:
        run_stats = new_run_stats()
    # 파일 읽기/결과 폴더 생성 단계에서 실패해도 XLSX 분석 1건의 실패로 집계.
    run_stats["attempted"] += 1
    workbook_completed = False
    temp_path = None
    try:
        if source.suffix.lower() not in TABLE_EXTENSIONS:
            raise ValueError("table_analysis는 CSV와 XLSX만 지원합니다.")
        relative_path = source.relative_to(INPUT_DIR)
        destination = OUTPUT_DIR / relative_path
        destination.mkdir(parents=True, exist_ok=True)
        result_path = destination / "컬럼분석.xlsx"
        charts_root = destination / "컬럼차트"
        with tempfile.NamedTemporaryFile(
            dir=destination, prefix="_analysis_", suffix=".xlsx", delete=False
        ) as stream:
            temp_path = Path(stream.name)
        sheet_count, chart_failures = 0, 0
        generated = []
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            for sheet_name, table in source_tables(source):
                report = analyze_columns(table)
                report.to_excel(writer, sheet_name=sheet_name, index=False)
                format_sheet(writer.sheets[sheet_name])
                sheet_count += 1
                chart_dir = charts_root / f"sheet_{sheet_count:03d}_{safe_filename(sheet_name)}"
                images, failed = column_charts(
                    table, sheet_name, chart_dir, chart_option, source, run_stats
                )
                generated.extend(images)
                chart_failures += failed
        temp_path.replace(result_path)
        workbook_completed = True
        run_stats["completed"] += 1
        if not chart_failures:
            try:
                remove_old_charts(charts_root, generated)
            except Exception as exc:
                # 생성 성공 건수와 과거 생성물 정리 실패는 구분합니다.
                print(f"[이전 차트 정리 실패] {source.name}: {exc}")
                log_error(exc, {"source_file": str(source)}, stage="chart_cleanup")
        return result_path, sheet_count, len(generated), chart_failures
    except BaseException:
        if not workbook_completed:
            run_stats["errors"] += 1
        raise
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def main() -> int:
    reset_error_file_limit()
    chart_option = choose_chart_option()
    if chart_option is None:
        return 0
    print(f"선택한 차트: {CHART_OPTIONS[chart_option]['label']}")
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    paths = sorted(
        INPUT_DIR.rglob("*"),
        key=lambda path: path.relative_to(INPUT_DIR).as_posix().casefold(),
    )
    all_sources = [
        path for path in paths
        if path.is_file() and not path.name.startswith("~$")
    ]
    # 미지원 원본도 요약에 포함하되 실행/완료/에러는 0으로 기록합니다.
    stats_by_source = {path: new_run_stats() for path in all_sources}
    exit_code = 0
    summary_failed = False
    try:
        for path in paths:
            if path.is_dir():
                (OUTPUT_DIR / path.relative_to(INPUT_DIR)).mkdir(
                    parents=True, exist_ok=True
                )
        targets = [
            path for path in all_sources
            if path.suffix.lower() in TABLE_EXTENSIONS
        ]
        skipped_files = [
            path for path in paths
            if path.is_file() and (
                path.suffix.lower() not in TABLE_EXTENSIONS
                or path.name.startswith("~$")
            )
        ]
        unsupported_counts, temporary_counts = Counter(), Counter()
        for path in skipped_files:
            extension = path.suffix.lower() or "(확장자 없음)"
            if path.name.startswith("~$"):
                temporary_counts[extension] += 1
            else:
                unsupported_counts[extension] += 1
        print(f"대상 폴더: {INPUT_DIR}")
        print(f"출력 폴더: {OUTPUT_DIR}")
        print(f"분석 대상 {len(targets)}개 / 미지원·임시 파일 건너뜀 {len(skipped_files)}개")
        for extension in sorted(set(unsupported_counts) | set(temporary_counts)):
            unsupported = unsupported_counts[extension]
            temporary = temporary_counts[extension]
            reasons = []
            if unsupported:
                reasons.append(f"미지원 {unsupported}개")
            if temporary:
                reasons.append(f"임시 {temporary}개")
            print(f"  - {extension}: {unsupported + temporary}개 ({', '.join(reasons)})")

        ready = False
        if not targets:
            print("분석할 CSV 또는 XLSX 파일이 없습니다.")
        else:
            try:
                if _COMMON_IMPORT_ERROR is not None:
                    raise _COMMON_IMPORT_ERROR
                configure_chart_backend(chart_option)
                ready = True
            except Exception as exc:
                log_error(exc, {
                    "option": chart_option, "chart_engine": CHART_OPTIONS[chart_option]["label"],
                    "python_executable": sys.executable, "python_version": sys.version,
                }, stage="dependency_setup")
                packages = {"1": "plotly", "2": "plotly kaleido", "3": "matplotlib"}[chart_option]
                print(f"[실행 준비 실패] {exc}")
                print(f"사용할 가상환경의 필요 패키지: pandas openpyxl numpy {packages}")
                exit_code = 1

        success = failure = 0
        if ready:
            for source in targets:
                relative_path = source.relative_to(INPUT_DIR)
                try:
                    output, sheet_count, images, failed = table_analysis(
                        source, chart_option, stats_by_source[source]
                    )
                    print(
                        f"[완료] {relative_path} ({sheet_count}개 시트, "
                        f"차트 {images}개 / 차트 실패 {failed}개) -> {output}"
                    )
                    success += 1
                except Exception as exc:
                    print(f"[실패] {relative_path}: {type(exc).__name__}: {exc}")
                    log_error(exc, {
                        "source_file": str(source), "option": chart_option,
                        "chart_engine": CHART_OPTIONS[chart_option]["label"],
                    }, stage="file_analysis")
                    failure += 1
        chart_total = sum(stats["chart_completed"] for stats in stats_by_source.values())
        chart_failures = sum(stats["chart_errors"] for stats in stats_by_source.values())
        print(f"\n처리 결과: 성공 {success}개 / 실패 {failure}개")
        print(f"차트 결과: 생성 {chart_total}개 / 실패 {chart_failures}개")
        if failure or chart_failures:
            exit_code = 1
    finally:
        # 일부 파일 실패/사용자 중단 시에도 현재까지의 집계를 저장합니다.
        try:
            summary_path = save_execution_summary(stats_by_source)
            print(f"[실행 요약 저장] {summary_path}")
        except Exception as exc:
            summary_failed = True
            log_error(exc, {"output_file": str(OUTPUT_DIR / "추가분석_실행요약.xlsx")},
                      stage="summary_write")
            print(f"[실행 요약 저장 실패] {exc} (Excel에서 열려 있다면 닫아 주세요.)")
    return 1 if summary_failed else exit_code


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n사용자가 작업을 중단했습니다.")
        sys.exit(130)
    except Exception as exc:
        log_error(exc, {"base_dir": str(BASE_DIR)}, stage="program")
        print(f"[실행 실패] {type(exc).__name__}: {exc}")
        sys.exit(1)
