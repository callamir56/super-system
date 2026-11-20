"""ساخت شکل‌های پژوهش (نمودارها و طرح‌واره‌ها) با برچسب فارسی."""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(BASE), "assets", "figures")
os.makedirs(OUT, exist_ok=True)

for f in ("Vazirmatn-Regular.ttf", "Vazirmatn-Bold.ttf", "Vazirmatn-Medium.ttf"):
    fm.fontManager.addfont(os.path.join(BASE, f))
plt.rcParams.update({
    "font.family": "Vazirmatn",
    "figure.dpi": 200,
    "savefig.dpi": 200,
    "axes.edgecolor": "#8A94A6",
    "text.color": "#2D3440",
    "axes.labelcolor": "#2D3440",
    "xtick.color": "#2D3440",
    "ytick.color": "#2D3440",
})
NAVY, TEAL, GRAY, LIGHT = "#1C375B", "#2A8297", "#2D3440", "#E8EEF4"
_FA = str.maketrans("0123456789.", "۰۱۲۳۴۵۶۷۸۹٫")


def fa_num(x):
    return str(x).translate(_FA)


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.savefig(p, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig)
    print("saved", p)


# ---------------------------------------------------------------- شکل ۱: قلب
def fig_heart():
    fig, ax = plt.subplots(figsize=(7.4, 4.8))
    ax.set_xlim(0, 10); ax.set_ylim(0, 7.0); ax.axis("off")

    def box(x, y, w, h, label, sub, fc):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05,rounding_size=0.18",
                                    fc=fc, ec=NAVY, lw=1.4))
        ax.text(x + w / 2, y + h * 0.63, label, ha="center", va="center", fontsize=11, weight="bold", color=NAVY)
        ax.text(x + w / 2, y + h * 0.27, sub, ha="center", va="center", fontsize=8.5, color=GRAY)

    box(5.60, 4.15, 3.00, 1.45, "دهلیز راست", "دریافت خون کم‌اکسیژن", "#DCE9F2")
    box(5.60, 1.75, 3.00, 1.55, "بطن راست", "پمپ به ریه", "#DCE9F2")
    box(1.40, 4.15, 3.00, 1.45, "دهلیز چپ", "دریافت خون از ریه", "#F3E2DC")
    box(1.40, 1.75, 3.00, 1.55, "بطن چپ", "پمپ به کل بدن", "#F3E2DC")

    # دریچه‌ها
    ax.plot([5.60, 8.60], [3.72, 3.72], color=TEAL, lw=2.6, solid_capstyle="butt")
    ax.plot([1.40, 4.40], [3.72, 3.72], color=TEAL, lw=2.6, solid_capstyle="butt")
    ax.text(8.50, 3.80, "دریچه سهلتی", ha="right", va="bottom", fontsize=8.5, color=TEAL)
    ax.text(1.50, 3.80, "دریچه میترال", ha="left", va="bottom", fontsize=8.5, color=TEAL)

    # جریان خون
    ax.add_patch(FancyArrowPatch((7.10, 6.40), (7.10, 5.62), arrowstyle="-|>", mutation_scale=15, color="#C0392B", lw=1.9))
    ax.text(7.28, 6.02, "از بدن (سیاهرگ اجوف)", fontsize=8, color="#C0392B", va="center")
    ax.add_patch(FancyArrowPatch((2.90, 5.62), (2.90, 6.40), arrowstyle="-|>", mutation_scale=15, color="#C0392B", lw=1.9))
    ax.text(3.08, 6.02, "از ریه (سیاهرگ ریوی)", fontsize=8, color="#C0392B", va="center")
    ax.add_patch(FancyArrowPatch((8.72, 2.52), (9.80, 2.52), arrowstyle="-|>", mutation_scale=15, color="#2A6F97", lw=1.9))
    ax.text(9.26, 2.88, "به ریه", fontsize=8, color="#2A6F97", ha="center")
    ax.add_patch(FancyArrowPatch((1.28, 2.52), (0.20, 2.52), arrowstyle="-|>", mutation_scale=17, color="#C0392B", lw=2.8))
    ax.text(0.74, 2.88, "آئورت (به بدن)", fontsize=8, color="#C0392B", ha="center")

    # دیواره ضخیم بطن چپ
    ax.plot([1.40, 4.40], [1.92, 1.92], color=NAVY, lw=3.8, solid_capstyle="butt")
    ax.text(2.90, 1.48, "دیواره عضلانی ضخیم‌تر", fontsize=8, color=NAVY, ha="center")

    ax.text(5.0, 0.72, "طرح‌واره ساده‌شده قلب: چهار حفره، چهار دریچه و مسیر جریان خون",
            ha="center", fontsize=9.5, color="#5A6472", style="italic")
    save(fig, "fig_heart.png")


# ---------------------------------------------------------------- شکل ۲: نوار قلب
def _beat(t, rate=1.0, p=True):
    y = np.zeros_like(t)
    for c in np.arange(0.4, t[-1], rate):
        if p:
            y += 0.16 * np.exp(-((t - (c - 0.20)) ** 2) / (2 * 0.016 ** 2))
        y += -0.10 * np.exp(-((t - (c - 0.025)) ** 2) / (2 * 0.008 ** 2))
        y += 1.00 * np.exp(-((t - c) ** 2) / (2 * 0.006 ** 2))
        y += -0.22 * np.exp(-((t - (c + 0.028)) ** 2) / (2 * 0.009 ** 2))
        y += 0.28 * np.exp(-((t - (c + 0.22)) ** 2) / (2 * 0.045 ** 2))
    return y


def fig_ecg():
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(7.4, 4.2), sharex=True)
    t = np.linspace(0, 4, 4000)

    a1.plot(t, _beat(t, 0.80), color=NAVY, lw=1.3)
    a1.set_title("ریتم سینوسی طبیعی: موج P پیش از هر کمپلکس QRS", fontsize=10.5, color=NAVY, pad=6)
    # برچسب‌گذاری روی دومین ضربان (مرکز ۱٫۲ ثانیه)
    for x, ytxt, xt, yt, ha, lab in ((1.00, 0.30, 0.68, 1.34, "center", "موج P"),
                                     (1.20, 1.06, 1.24, 1.56, "center", "کمپلکس QRS"),
                                     (1.42, 0.34, 1.80, 1.34, "center", "موج T")):
        a1.annotate(lab, xy=(x, ytxt), xytext=(xt, yt), ha=ha, va="center", fontsize=8.5,
                    color=TEAL, arrowprops=dict(arrowstyle="-", color=TEAL, lw=0.9))
    a1.annotate("", xy=(1.27, 0.68), xytext=(2.00, 0.68),
                arrowprops=dict(arrowstyle="<|-|>", color="#C0392B", lw=1.1))
    a1.text(1.635, 0.76, "فاصله RR", ha="center", fontsize=8, color="#C0392B",
            bbox=dict(fc="white", ec="none", pad=0.6))

    # فیبریلاسیون دهلیزی: فواصل نامنظم و بدون موج P
    t2 = np.linspace(0, 4, 4000)
    y2 = np.zeros_like(t2)
    rng = np.random.default_rng(7)
    times, cur = [], 0.35
    while cur < 3.8:
        times.append(cur)
        cur += rng.uniform(0.30, 1.15)
    for c in times:
        y2 += -0.10 * np.exp(-((t2 - (c - 0.025)) ** 2) / (2 * 0.008 ** 2))
        y2 += 1.00 * np.exp(-((t2 - c) ** 2) / (2 * 0.006 ** 2))
        y2 += -0.22 * np.exp(-((t2 - (c + 0.028)) ** 2) / (2 * 0.009 ** 2))
        y2 += 0.26 * np.exp(-((t2 - (c + 0.22)) ** 2) / (2 * 0.045 ** 2))
    y2 += 0.045 * np.sin(2 * np.pi * 6.5 * t2) + 0.03 * np.sin(2 * np.pi * 9.3 * t2 + 1.1)
    a2.plot(t2, y2, color="#C0392B", lw=1.2)
    a2.set_title("فیبریلاسیون دهلیزی: بدون موج P و با فواصل RR نامنظم", fontsize=10.5, color="#C0392B", pad=6)

    for ax in (a1, a2):
        ax.set_yticks([]); ax.set_ylim(-0.45, 1.78)
        ax.set_xticks([])
        for sp in ("top", "right", "left", "bottom"):
            ax.spines[sp].set_visible(False)
    a2.set_xlabel("زمان (ثانیه)", fontsize=9)
    fig.text(0.5, -0.05, "سیگنال‌های بالا نمونه‌های شبیه‌سازی‌شده برای آموزش مفهوم‌اند و ثبت واقعی بیمار نیستند.",
             ha="center", fontsize=8.5, color="#6B7280", style="italic")
    fig.tight_layout(h_pad=1.6)
    save(fig, "fig_ecg.png")


# ---------------------------------------------------------------- شکل ۳: فلوچارت ساخت مدل
def fig_pipeline():
    steps = ["تعریف دقیق پرسش بالینی", "جمع‌آوری داده از منابع معتبر",
             "پاک‌سازی، یکسان‌سازی و برچسب‌گذاری", "تقسیم به آموزش، اعتبارسنجی و آزمون",
             "آموزش مدل و تنظیم پارامترها", "سنجش روی داده آزمون مستقل",
             "ارزیابی در محیط واقعی و بررسی سوگیری", "استقرار همراه با نظارت انسانی و پایش"]
    fig, ax = plt.subplots(figsize=(6.6, 6.4))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    h, gap = 0.86, 0.34
    for i, s in enumerate(steps):
        y = 9.1 - i * (h + gap)
        fc = LIGHT if i % 2 == 0 else "#DCE9F2"
        ax.add_patch(FancyBboxPatch((0.6, y), 8.8, h, boxstyle="round,pad=0.03,rounding_size=0.12",
                                    fc=fc, ec=NAVY, lw=1.2))
        ax.text(5.0, y + h / 2, f"گام {fa_num(i+1)}: {s}", ha="center", va="center", fontsize=9.5, color=NAVY)
        if i < len(steps) - 1:
            ax.add_patch(FancyArrowPatch((5.0, y - 0.03), (5.0, y - gap + 0.03),
                                         arrowstyle="-|>", mutation_scale=13, color=TEAL, lw=1.5))
    ax.text(5.0, 0.15, "بازگشت به مرحله داده در صورت افت عملکرد در محیط واقعی",
            ha="center", fontsize=8.5, color="#C0392B", style="italic")
    save(fig, "fig_pipeline.png")


# ---------------------------------------------------------------- شکل ۴: ماتریس درهم‌ریختگی
def fig_confusion(tp=90, fn=10, fp=90, tn=810):
    fig, ax = plt.subplots(figsize=(5.8, 4.1))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    # چیدمان راست‌به‌چپ: ستون «واقعاً بیمار» در سمت راست
    cells = [(5.4, 5.1, tp, "مثبت درست\n(TP)", "#CDE3D6"),
             (1.0, 5.1, fn, "منفی کاذب\n(FN)", "#F6D5CF"),
             (5.4, 1.6, fp, "مثبت کاذب\n(FP)", "#F6D5CF"),
             (1.0, 1.6, tn, "منفی درست\n(TN)", "#CDE3D6")]
    for x, y, v, lab, c in cells:
        ax.add_patch(Rectangle((x, y), 3.9, 3.2, fc=c, ec="white", lw=2))
        ax.text(x + 1.95, y + 2.15, fa_num(v), ha="center", va="center", fontsize=20, weight="bold", color=NAVY)
        ax.text(x + 1.95, y + 0.85, lab, ha="center", va="center", fontsize=9, color=GRAY)
    ax.text(7.35, 9.25, "واقعاً بیمار", ha="center", fontsize=11, weight="bold", color=NAVY)
    ax.text(2.95, 9.25, "واقعاً سالم", ha="center", fontsize=11, weight="bold", color=NAVY)
    ax.text(0.45, 6.7, "مدل: بیمار", ha="center", va="center", fontsize=11, weight="bold", color=NAVY, rotation=90)
    ax.text(0.45, 3.2, "مدل: سالم", ha="center", va="center", fontsize=11, weight="bold", color=NAVY, rotation=90)
    ax.text(5.15, 0.55, "ماتریس درهم‌ریختگی در یک مثال فرضی: ۱۰۰۰ بیمار، شیوع ۱۰٪",
            ha="center", fontsize=8.5, color="#6B7280", style="italic")
    save(fig, "fig_confusion.png")


# ---------------------------------------------------------------- شکل ۵: ROC
def fig_roc():
    fig, ax = plt.subplots(figsize=(5.4, 4.4))
    fpr = np.linspace(0, 1, 400)

    for target, lab, c, ls in ((0.90, "مدل الف", NAVY, "-"),
                               (0.77, "مدل ب", TEAL, "--")):
        a = 1 / target - 1
        tpr = fpr ** a
        auc = np.trapezoid(tpr, fpr)
        ax.plot(fpr, tpr, color=c, lw=2, ls=ls,
                label=f"{lab} (AUC ≈ {fa_num(round(auc, 2))})")
    ax.plot([0, 1], [0, 1], color="#9AA3AE", lw=1, ls=":", label="پیش‌بینی تصادفی (AUC = ۰٫۵)")
    ax.set_xlabel("نرخ مثبت کاذب (۱ − ویژگی)", fontsize=9.5)
    ax.set_ylabel("نرخ مثبت درست (حساسیت)", fontsize=9.5)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ticks = [0, 0.2, 0.4, 0.6, 0.8, 1.0]
    ax.set_xticks(ticks); ax.set_yticks(ticks)
    ax.set_xticklabels([fa_num(f"{v:.1f}") for v in ticks], fontsize=8.5)
    ax.set_yticklabels([fa_num(f"{v:.1f}") for v in ticks], fontsize=8.5)
    ax.grid(color="#E2E8EF", lw=0.6)
    ax.legend(loc="lower right", fontsize=8.5, frameon=False)
    ax.set_title("منحنی ROC: هرچه منحنی به گوشه بالا-چپ نزدیک‌تر، تمایز بهتر", fontsize=9.5, color=NAVY, pad=8)
    fig.text(0.5, -0.10, "منحنی‌ها نمونه‌ای و برای آموزش مفهوم ترسیم شده‌اند، نه نتیجه یک مطالعه واقعی.",
             ha="center", fontsize=8, color="#6B7280", style="italic")
    save(fig, "fig_roc.png")


# ---------------------------------------------------------------- شکل ۶: اثر شیوع بر PPV
def ppv(sens, spec, prev):
    tp = sens * prev
    fp = (1 - spec) * (1 - prev)
    return tp / (tp + fp) if (tp + fp) else 0


def fig_prevalence():
    prevs = np.array([1, 2, 5, 10, 20, 30])
    s1 = [ppv(0.90, 0.90, p / 100) * 100 for p in prevs]
    s2 = [ppv(0.90, 0.95, p / 100) * 100 for p in prevs]
    for p, a, b in zip(prevs, s1, s2):
        print(f"prev {p}%: PPV spec90={a:.1f}%  spec95={b:.1f}%")

    fig, ax = plt.subplots(figsize=(6.4, 3.8))
    x = np.arange(len(prevs)); w = 0.38
    ax.bar(x - w / 2, s1, w, color=TEAL, label="ویژگی ۹۰٪")
    ax.bar(x + w / 2, s2, w, color=NAVY, label="ویژگی ۹۵٪")
    for xi, a, b in zip(x, s1, s2):
        ax.text(xi - w / 2, a + 2.0, fa_num(f"{a:.0f}"), ha="center", fontsize=8, color=TEAL)
        ax.text(xi + w / 2, b + 2.0, fa_num(f"{b:.0f}"), ha="center", fontsize=8, color=NAVY)
    ax.set_xticks(x); ax.set_xticklabels([fa_num(p) for p in prevs], fontsize=9)
    ax.set_yticks([0, 20, 40, 60, 80, 100])
    ax.set_yticklabels([fa_num(v) for v in [0, 20, 40, 60, 80, 100]], fontsize=8.5)
    ax.set_xlabel("شیوع بیماری در جمعیت (درصد)", fontsize=9.5)
    ax.set_ylabel("ارزش اخباری مثبت (درصد)", fontsize=9.5)
    ax.set_ylim(0, 108)
    ax.grid(axis="y", color="#E2E8EF", lw=0.6)
    ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.legend(fontsize=8.5, frameon=False, loc="upper left")
    ax.set_title("حساسیت ثابت ۹۰٪: با کاهش شیوع، اعتبار هشدار مثبت افت می‌کند", fontsize=9.5, color=NAVY, pad=8)
    save(fig, "fig_prevalence.png")


if __name__ == "__main__":
    fig_heart()
    fig_ecg()
    fig_pipeline()
    fig_confusion()
    fig_roc()
    fig_prevalence()
