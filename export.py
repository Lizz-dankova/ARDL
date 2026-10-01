# ============================================================
# ARDL ECONOMETRIC ANALYSIS
# ============================================================
#
# Dataset:
#   2000–2024
#
# Excel:
#   sheet = "stat"
#   C:\Users\zinag\Documents\statistics.xlsx
# Variables:
#   Year
#   Government expenditure on education, total (% of GDP)
#   GERD as a percentage of GDP
#   Information technology exports, percent of total goods exports
#
# Dependent variable:
#   IT_exports
#
# Explanatory variables:
#   Education
#   GERD
#
# ============================================================

import os
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
import statsmodels.api as sm

matplotlib.use("Agg")

import matplotlib.pyplot as plt

from scipy import stats
from statsmodels.tsa.stattools import adfuller, kpss, acf, pacf
from statsmodels.tsa.stattools import zivot_andrews
from statsmodels.tsa.stattools import arma_order_select_ic

from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import (
    acorr_breusch_godfrey,
    het_breuschpagan,
    het_white,
    linear_reset,
)
from statsmodels.stats.stattools import jarque_bera

from statsmodels.tsa.ardl import ARDL, UECM

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# ДОПОМІЖНА ФУНКЦІЯ: ADJUSTED R²
# ============================================================

def calculate_adjusted_r2(result):

    r2 = result.rsquared
    n = result.nobs
    k = len(result.params) - 1

    denominator = n - k - 1

    if denominator <= 0:
        return np.nan

    return 1 - (1 - r2) * (n - 1) / denominator


def print_final_ardl_equation(result, p, q_education, q_gerd):
    """
    Виведення остаточної числової ARDL-моделі
    разом із p-value, R² та іншими статистиками.
    """

    params = result.params

    print("\n" + "=" * 80)
    print("ОСТАТОЧНА ЧИСЛОВА ARDL-МОДЕЛЬ")
    print("=" * 80)

    # --------------------------------------------------------
    # Загальний вигляд
    # --------------------------------------------------------
    print("\nЗагальний вигляд моделі:")
    print(
        f"IT_exports_t = "
        f"f(IT_exports_(t-1), Education_(t-{q_education}), "
        f"GERD_(t-{q_gerd}))"
    )

    # --------------------------------------------------------
    # Числова модель
    # --------------------------------------------------------
    print("\nЧислове рівняння:")

    equation = f"IT_exports_t = {params['const']:.6f}"

    # Лаги залежної змінної
    for lag in range(1, p + 1):
        name = f"IT_exports.L{lag}"

        if name in params.index:
            coef = params[name]

            if coef >= 0:
                equation += f" + {coef:.6f}·IT_exports_(t-{lag})"
            else:
                equation += f" - {abs(coef):.6f}·IT_exports_(t-{lag})"

    # Education
    for lag in range(0, q_education + 1):
        name = f"Education.L{lag}"

        if name in params.index:
            coef = params[name]

            if lag == 0:
                variable = "Education_t"
            else:
                variable = f"Education_(t-{lag})"

            if coef >= 0:
                equation += f" + {coef:.6f}·{variable}"
            else:
                equation += f" - {abs(coef):.6f}·{variable}"

    # GERD
    for lag in range(0, q_gerd + 1):
        name = f"GERD.L{lag}"

        if name in params.index:
            coef = params[name]

            if lag == 0:
                variable = "GERD_t"
            else:
                variable = f"GERD_(t-{lag})"

            if coef >= 0:
                equation += f" + {coef:.6f}·{variable}"
            else:
                equation += f" - {abs(coef):.6f}·{variable}"

    equation += " + ε_t"

    print(equation)

    # --------------------------------------------------------
    # Таблиця коефіцієнтів
    # --------------------------------------------------------
    print("\n" + "-" * 80)
    print("ОЦІНКА КОЕФІЦІЄНТІВ")
    print("-" * 80)

    coef_table = pd.DataFrame({
        "Coefficient": result.params,
        "Std. Error": result.bse,
        "t-statistic": result.tvalues,
        "p-value": result.pvalues
    })

    print(coef_table.to_string(float_format=lambda x: f"{x:.6f}"))

    # --------------------------------------------------------
    # Основні статистики моделі
    # --------------------------------------------------------
    print("\n" + "-" * 80)
    print("ОСНОВНІ СТАТИСТИКИ ОСТАТОЧНОЇ МОДЕЛІ")
    print("-" * 80)

    r2 = result.rsquared

    # Adjusted R²
    n = result.nobs
    k = len(result.params) - 1

    if n - k - 1 > 0:
        adjusted_r2 = 1 - (1 - r2) * (n - 1) / (n - k - 1)
    else:
        adjusted_r2 = np.nan

    print(f"Number of observations : {n}")
    print(f"R²                     : {r2:.6f}")
    print(f"Adjusted R²            : {adjusted_r2:.6f}")
    print(f"Log-Likelihood         : {result.llf:.6f}")
    print(f"AIC                    : {result.aic:.6f}")
    print(f"BIC                    : {result.bic:.6f}")
    print(f"HQIC                   : {result.hqic:.6f}")

    # --------------------------------------------------------
    # Значущість коефіцієнтів
    # --------------------------------------------------------
    print("\n" + "-" * 80)
    print("ІНТЕРПРЕТАЦІЯ p-value")
    print("-" * 80)

    for name, pvalue in result.pvalues.items():

        if pvalue < 0.01:
            significance = "статистично значущий на рівні 1%"
        elif pvalue < 0.05:
            significance = "статистично значущий на рівні 5%"
        elif pvalue < 0.10:
            significance = "статистично значущий на рівні 10%"
        else:
            significance = "статистично незначущий"

        print(
            f"{name}: p-value = {pvalue:.6f} → {significance}"
        )

    print("=" * 80)


# ============================================================
# 1. ЗАВАНТАЖЕННЯ ДАНИХ
# ============================================================

def load_excel_data():

    print("=" * 100)
    print("1. ЗАВАНТАЖЕННЯ ДАНИХ")
    print("=" * 100)

    path = input("Введіть шлях до Excel-файлу: ").strip()

    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Файл не знайдено:\n{path}"
        )

    print("\nЧитання аркуша 'stat'...")

    data = pd.read_excel(
        path,
        sheet_name="stat"
    )

    print(f"Початкова кількість рядків: {len(data)}")
    print(f"Початкові назви стовпців:\n{list(data.columns)}")

    # --------------------------------------------------------
    # Перейменування
    # --------------------------------------------------------

    rename_dict = {
        "Government expenditure on education, total (% of GDP)": "Education",
        "GERD as a percentage of GDP": "GERD",
        "Information technology exports, percent of total goods exports": "IT_exports"
    }

    data = data.rename(columns=rename_dict)

    required_columns = [
        "Year",
        "Education",
        "GERD",
        "IT_exports"
    ]

    missing = [
        col for col in required_columns
        if col not in data.columns
    ]

    if missing:
        raise ValueError(
            f"\nВідсутні необхідні стовпці:\n{missing}"
        )

    # --------------------------------------------------------
    # Перетворення типів
    # --------------------------------------------------------

    for col in required_columns:
        data[col] = pd.to_numeric(
            data[col],
            errors="coerce"
        )

    # --------------------------------------------------------
    # Фільтрація 2000–2024
    # --------------------------------------------------------

    data = data[
        (data["Year"] >= 2000) &
        (data["Year"] <= 2024)
    ].copy()

    data = data.sort_values("Year")
    data = data.reset_index(drop=True)

    print("\nПеріод аналізу:")
    print(
        f"{int(data['Year'].min())}–"
        f"{int(data['Year'].max())}"
    )

    print(f"Кількість спостережень: {len(data)}")

    # --------------------------------------------------------
    # Перевірка пропусків
    # --------------------------------------------------------

    print("\nПропущені значення:")

    print(
        data[
            required_columns
        ].isna().sum()
    )

    if data[required_columns].isna().any().any():

        print(
            "\nУВАГА: знайдено пропущені значення."
        )

        data = data.dropna(
            subset=required_columns
        ).copy()

    # --------------------------------------------------------
    # Перевірка років
    # --------------------------------------------------------

    expected_years = set(
        range(2000, 2025)
    )

    actual_years = set(
        data["Year"].astype(int)
    )

    missing_years = sorted(
        expected_years - actual_years
    )

    if missing_years:

        print(
            "\nВідсутні роки:"
        )
        print(missing_years)

    else:

        print(
            "\nУсі роки 2000–2024 присутні."
        )

    print("\nПідсумкова структура даних:")

    print(
        data[
            [
                "Year",
                "Education",
                "GERD",
                "IT_exports"
            ]
        ].to_string(index=False)
    )

    return data


# ============================================================
# BREUSCH-GODFREY ДЛЯ ARDL
# ============================================================

def breusch_godfrey_ardl(
    result,
    lag
):

    resid = pd.Series(
        result.resid
    ).dropna()

    # Матриця регресорів ARDL
    X = pd.DataFrame(
        result.model._x,
        index=resid.index
    )

    X = sm.add_constant(
        X,
        has_constant="add"
    )

    # Залежна змінна:
    # залишок поточного періоду
    y_aux = resid.copy()

    # Додаємо лаги залишків
    for i in range(1, lag + 1):

        X[
            f"resid_L{i}"
        ] = resid.shift(i)

    aux = pd.concat(
        [
            y_aux.rename("resid"),
            X
        ],
        axis=1
    ).dropna()

    y_aux = aux["resid"]

    X_aux = aux.drop(
        columns="resid"
    )

    auxiliary_model = sm.OLS(
        y_aux,
        X_aux
    ).fit()

    # LM statistic
    n = len(auxiliary_model.resid)

    R2 = auxiliary_model.rsquared

    LM = n * R2

    p_value = stats.chi2.sf(
        LM,
        lag
    )

    return (
        LM,
        p_value,
        auxiliary_model
    )


# ============================================================
# 2. ОПИСОВА СТАТИСТИКА
# ============================================================

def descriptive_statistics(data):

    print("\n" + "=" * 100)
    print("2. ОПИСОВА СТАТИСТИКА")
    print("=" * 100)

    variables = [
        "Education",
        "GERD",
        "IT_exports"
    ]

    desc = data[variables].describe().T

    desc["median"] = data[variables].median()

    desc = desc[
        [
            "count",
            "mean",
            "std",
            "min",
            "25%",
            "median",
            "50%",
            "75%",
            "max"
        ]
    ]

    print(desc)

    return desc


# ============================================================
# 3. ГРАФІКИ ЧАСОВИХ РЯДІВ
# ============================================================

def plot_time_series(data):

    print("\n" + "=" * 100)
    print("3. ГРАФІКИ ЧАСОВИХ РЯДІВ")
    print("=" * 100)

    variables = [
        "Education",
        "GERD",
        "IT_exports"
    ]

    os.makedirs("ARDL_results", exist_ok=True)

    for variable in variables:

        plt.figure(figsize=(11, 5))

        plt.plot(
            data["Year"],
            data[variable],
            marker="o"
        )

        plt.xlabel("Year")
        plt.ylabel(variable)
        plt.title(
            f"{variable}: 2000–2024"
        )

        plt.grid(
            True,
            alpha=0.3
        )

        plt.tight_layout()

        filename = (
            f"ARDL_results/"
            f"{variable}_time_series.png"
        )

        plt.savefig(
            filename,
            dpi=300
        )

        plt.close()

        print(
            f"Збережено: {filename}"
        )


# ============================================================
# 4. ТЕСТУВАННЯ КОРЕЛЯЦІЇ
# ============================================================

def correlation_analysis(data):

    print("\n" + "=" * 100)
    print("4. ТЕСТУВАННЯ КОРЕЛЯЦІЇ")
    print("=" * 100)

    variables = [
        "IT_exports",
        "Education",
        "GERD"
    ]

    # --------------------------------------------------------
    # Кореляційна матриця
    # --------------------------------------------------------

    pearson_matrix = data[
        variables
    ].corr(method="pearson")

    spearman_matrix = data[
        variables
    ].corr(method="spearman")

    print("\nКоефіцієнти кореляції Пірсона:")

    print(
        pearson_matrix.round(4)
    )

    print("\nКоефіцієнти кореляції Спірмена:")

    print(
        spearman_matrix.round(4)
    )

    # --------------------------------------------------------
    # Тести значущості парних кореляцій
    # --------------------------------------------------------

    rows = []

    for i in range(len(variables)):

        for j in range(i + 1, len(variables)):

            var1 = variables[i]
            var2 = variables[j]

            x = data[var1]
            y = data[var2]

            pearson_r, pearson_p = stats.pearsonr(
                x,
                y
            )

            spearman_rho, spearman_p = stats.spearmanr(
                x,
                y
            )

            rows.append(
                {
                    "Variable_1": var1,
                    "Variable_2": var2,
                    "Pearson_r": pearson_r,
                    "Pearson_p_value": pearson_p,
                    "Spearman_rho": spearman_rho,
                    "Spearman_p_value": spearman_p
                }
            )

    correlation_tests = pd.DataFrame(
        rows
    )

    print(
        "\nТести значущості кореляції:"
    )

    print(
        correlation_tests.round(6).to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Інтерпретація
    # --------------------------------------------------------

    print("\nІнтерпретація:")

    for _, row in correlation_tests.iterrows():

        print(
            f"\n{row['Variable_1']} – "
            f"{row['Variable_2']}"
        )

        print(
            f"Pearson r = "
            f"{row['Pearson_r']:.4f}, "
            f"p-value = "
            f"{row['Pearson_p_value']:.6f}"
        )

        print(
            f"Spearman rho = "
            f"{row['Spearman_rho']:.4f}, "
            f"p-value = "
            f"{row['Spearman_p_value']:.6f}"
        )

        if row["Pearson_p_value"] < 0.05:

            print(
                "Кореляція Пірсона статистично "
                "значуща на рівні 5%."
            )

        else:

            print(
                "Кореляція Пірсона статистично "
                "не значуща на рівні 5%."
            )

    # --------------------------------------------------------
    # Графік кореляцій
    # --------------------------------------------------------

    plt.figure(figsize=(8, 6))

    plt.imshow(
        pearson_matrix,
        aspect="auto"
    )

    plt.colorbar(
        label="Pearson correlation"
    )

    plt.xticks(
        range(len(variables)),
        variables,
        rotation=45,
        ha="right"
    )

    plt.yticks(
        range(len(variables)),
        variables
    )

    plt.title(
        "Pearson correlation matrix"
    )

    plt.tight_layout()

    filename = (
        "ARDL_results/"
        "correlation_matrix.png"
    )

    plt.savefig(
        filename,
        dpi=300
    )

    plt.close()

    print(
        f"\nМатрицю кореляцій збережено: "
        f"{filename}"
    )

    return (
        pearson_matrix,
        spearman_matrix,
        correlation_tests
    )


# ============================================================
# 5. ADF
# ============================================================

def adf_test(series, name, regression="c"):

    series = series.dropna()

    result = adfuller(
        series,
        autolag="AIC",
        regression=regression
    )

    statistic = result[0]
    p_value = result[1]
    used_lag = result[2]
    nobs = result[3]

    return {
        "Variable": name,
        "Test": "ADF",
        "Regression": regression,
        "Statistic": statistic,
        "p_value": p_value,
        "Lags": used_lag,
        "N_obs": nobs
    }


# ============================================================
# 6. PP
# ============================================================

def pp_test(series, name):

    try:

        from arch.unitroot import PhillipsPerron

        series = series.dropna()

        test = PhillipsPerron(
            series
        )

        return {
            "Variable": name,
            "Test": "PP",
            "Statistic": test.stat,
            "p_value": test.pvalue,
            "Lags": test.lags,
            "N_obs": len(series)
        }

    except Exception as e:

        return {
            "Variable": name,
            "Test": "PP",
            "Statistic": np.nan,
            "p_value": np.nan,
            "Lags": np.nan,
            "N_obs": len(series),
            "Error": str(e)
        }


# ============================================================
# 7. KPSS
# ============================================================

def kpss_test(series, name, regression="c"):

    series = series.dropna()

    try:

        statistic, p_value, lags, critical = kpss(
            series,
            regression=regression,
            nlags="auto"
        )

        return {
            "Variable": name,
            "Test": "KPSS",
            "Regression": regression,
            "Statistic": statistic,
            "p_value": p_value,
            "Lags": lags,
            "N_obs": len(series)
        }

    except Exception as e:

        return {
            "Variable": name,
            "Test": "KPSS",
            "Regression": regression,
            "Statistic": np.nan,
            "p_value": np.nan,
            "Lags": np.nan,
            "N_obs": len(series),
            "Error": str(e)
        }


# ============================================================
# 8. ZIVOT–ANDREWS
# ============================================================

def za_test(series, name):

    series = series.dropna()

    try:

        result = zivot_andrews(
            series,
            maxlag=5,
            regression="c",
            autolag="AIC"
        )

        return {
            "Variable": name,
            "Test": "Zivot-Andrews",
            "Statistic": result[0],
            "p_value": result[1],
            "Lags": result[2],
            "Break": result[4],
            "N_obs": len(series)
        }

    except Exception as e:

        return {
            "Variable": name,
            "Test": "Zivot-Andrews",
            "Statistic": np.nan,
            "p_value": np.nan,
            "Lags": np.nan,
            "Break": np.nan,
            "N_obs": len(series),
            "Error": str(e)
        }


# ============================================================
# 9. UNIT ROOT TESTS
# ============================================================

def unit_root_tests(data):

    print("\n" + "=" * 100)
    print("5. ТЕСТИ НА СТАЦІОНАРНІСТЬ: РІВНІ")
    print("=" * 100)

    variables = [
        "IT_exports",
        "Education",
        "GERD"
    ]

    results = []

    for variable in variables:

        series = data[variable]

        results.append(
            adf_test(
                series,
                variable
            )
        )

        results.append(
            pp_test(
                series,
                variable
            )
        )

        results.append(
            kpss_test(
                series,
                variable
            )
        )

        results.append(
            za_test(
                series,
                variable
            )
        )

    results_df = pd.DataFrame(
        results
    )

    print(
        results_df.to_string(
            index=False
        )
    )

    return results_df


# ============================================================
# 10. ПЕРШІ РІЗНИЦІ
# ============================================================

def first_difference_tests(data):

    print("\n" + "=" * 100)
    print("6. ТЕСТИ НА СТАЦІОНАРНІСТЬ: ПЕРШІ РІЗНИЦІ")
    print("=" * 100)

    variables = [
        "IT_exports",
        "Education",
        "GERD"
    ]

    results = []

    for variable in variables:

        diff = data[variable].diff().dropna()

        results.append(
            adf_test(
                diff,
                variable,
                regression="c"
            )
        )

        results.append(
            pp_test(
                diff,
                variable
            )
        )

        results.append(
            kpss_test(
                diff,
                variable,
                regression="c"
            )
        )

        results.append(
            za_test(
                diff,
                variable
            )
        )

    results_df = pd.DataFrame(
        results
    )

    print(
        results_df.to_string(
            index=False
        )
    )

    return results_df


# ============================================================
# 11. ДРУГІ РІЗНИЦІ: ПЕРЕВІРКА I(2)
# ============================================================

def second_difference_tests(data):

    print("\n" + "=" * 100)
    print("7. ПЕРЕВІРКА НА I(2)")
    print("=" * 100)

    print(
        """
Для ARDL Bounds Testing необхідно, щоб змінні були
I(0) або I(1), але не I(2).

Тому додатково перевіряється стаціонарність
других різниць.

Якщо змінна є нестабільною у рівнях та перших різницях,
але стає стаціонарною після другого диференціювання,
це є ознакою I(2).
"""
    )

    variables = [
        "IT_exports",
        "Education",
        "GERD"
    ]

    results = []

    for variable in variables:

        second_diff = (
            data[variable]
            .diff()
            .diff()
            .dropna()
        )

        results.append(
            adf_test(
                second_diff,
                variable,
                regression="c"
            )
        )

        results.append(
            pp_test(
                second_diff,
                variable
            )
        )

        results.append(
            kpss_test(
                second_diff,
                variable,
                regression="c"
            )
        )

    results_df = pd.DataFrame(
        results
    )

    print(
        "\nРезультати тестів других різниць:"
    )

    print(
        results_df.to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Аналіз результатів
    # --------------------------------------------------------

    print("\nВисновок щодо I(2):")

    i2_summary = []

    for variable in variables:

        temp = results_df[
            results_df["Variable"] == variable
        ].copy()

        adf_p = temp.loc[
            temp["Test"] == "ADF",
            "p_value"
        ]

        pp_p = temp.loc[
            temp["Test"] == "PP",
            "p_value"
        ]

        kpss_p = temp.loc[
            temp["Test"] == "KPSS",
            "p_value"
        ]

        adf_p = (
            adf_p.iloc[0]
            if len(adf_p)
            else np.nan
        )

        pp_p = (
            pp_p.iloc[0]
            if len(pp_p)
            else np.nan
        )

        kpss_p = (
            kpss_p.iloc[0]
            if len(kpss_p)
            else np.nan
        )

        stationary_votes = 0

        # ADF: H0 = unit root
        if (
            not pd.isna(adf_p)
            and adf_p < 0.05
        ):
            stationary_votes += 1

        # PP: H0 = unit root
        if (
            not pd.isna(pp_p)
            and pp_p < 0.05
        ):
            stationary_votes += 1

        # KPSS: H0 = stationarity
        if (
            not pd.isna(kpss_p)
            and kpss_p >= 0.05
        ):
            stationary_votes += 1

        if stationary_votes >= 2:

            conclusion = (
                "Другі різниці стаціонарні. "
                "Це є сумісним з I(2), якщо "
                "рівні та перші різниці "
                "відповідно були нестабільними."
            )

        else:

            conclusion = (
                "Переконливих ознак стаціонарності "
                "других різниць за більшістю тестів "
                "не отримано."
            )

        print(
            f"\n{variable}:"
        )

        print(
            f"ADF p-value = {adf_p}"
        )

        print(
            f"PP p-value = {pp_p}"
        )

        print(
            f"KPSS p-value = {kpss_p}"
        )

        print(
            conclusion
        )

        i2_summary.append(
            {
                "Variable": variable,
                "ADF_p_value_2diff": adf_p,
                "PP_p_value_2diff": pp_p,
                "KPSS_p_value_2diff": kpss_p,
                "Stationarity_votes": stationary_votes,
                "I2_warning": (
                    "YES"
                    if stationary_votes >= 2
                    else "NO"
                ),
                "Conclusion": conclusion
            }
        )

    summary_df = pd.DataFrame(
        i2_summary
    )

    return (
        results_df,
        summary_df
    )


# ============================================================
# 12. ПІДСУМКОВА КЛАСИФІКАЦІЯ I(0)/I(1)/I(2)
# ============================================================

def integration_summary(
    level_results,
    diff_results,
    second_diff_summary
):

    print("\n" + "=" * 100)
    print("8. ПІДСУМКОВА КЛАСИФІКАЦІЯ ПОРЯДКУ ІНТЕГРАЦІЇ")
    print("=" * 100)

    variables = [
        "IT_exports",
        "Education",
        "GERD"
    ]

    rows = []

    for variable in variables:

        level = level_results[
            level_results["Variable"] == variable
        ]

        diff = diff_results[
            diff_results["Variable"] == variable
        ]

        second = second_diff_summary[
            second_diff_summary["Variable"] == variable
        ]

        # ----------------------------------------------------
        # Рівень
        # ----------------------------------------------------

        level_stationary = 0

        for test in ["ADF", "PP", "KPSS"]:

            row = level[
                level["Test"] == test
            ]

            if row.empty:
                continue

            p = row["p_value"].iloc[0]

            if pd.isna(p):
                continue

            if test in ["ADF", "PP"]:

                if p < 0.05:
                    level_stationary += 1

            elif test == "KPSS":

                if p >= 0.05:
                    level_stationary += 1

        # ----------------------------------------------------
        # Перша різниця
        # ----------------------------------------------------

        diff_stationary = 0

        for test in ["ADF", "PP", "KPSS"]:

            row = diff[
                diff["Test"] == test
            ]

            if row.empty:
                continue

            p = row["p_value"].iloc[0]

            if pd.isna(p):
                continue

            if test in ["ADF", "PP"]:

                if p < 0.05:
                    diff_stationary += 1

            elif test == "KPSS":

                if p >= 0.05:
                    diff_stationary += 1

        # ----------------------------------------------------
        # Друга різниця
        # ----------------------------------------------------

        second_votes = (
            second["Stationarity_votes"].iloc[0]
            if not second.empty
            else np.nan
        )

        # ----------------------------------------------------
        # Класифікація
        # ----------------------------------------------------

        if level_stationary >= 2:

            integration_order = "I(0)"

        elif diff_stationary >= 2:

            integration_order = "I(1)"

        elif (
            not pd.isna(second_votes)
            and second_votes >= 2
        ):

            integration_order = (
                "Можлива I(2)"
            )

        else:

            integration_order = (
                "Невизначено"
            )

        rows.append(
            {
                "Variable": variable,
                "Level_stationary_tests": level_stationary,
                "First_difference_stationary_tests": diff_stationary,
                "Second_difference_stationary_tests": second_votes,
                "Integration_order": integration_order
            }
        )

    summary = pd.DataFrame(
        rows
    )

    print(
        summary.to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Загальний висновок
    # --------------------------------------------------------

    possible_i2 = summary[
        summary["Integration_order"]
        == "Можлива I(2)"
    ]

    print("\nЗагальний висновок:")

    if len(possible_i2) > 0:

        print(
            "УВАГА: для деяких змінних отримано "
            "ознаки можливої I(2)."
        )

        print(
            "Bounds Testing ARDL не слід інтерпретувати "
            "як стандартний тест коінтеграції, доки "
            "порядок інтеграції не буде обґрунтовано."
        )

    else:

        print(
            "За застосованим набором тестів "
            "переконливих ознак I(2) не виявлено."
        )

        print(
            "Це означає, що ARDL Bounds Testing "
            "можна продовжувати за умови, що "
            "остаточний висновок щодо порядку інтеграції "
            "ґрунтується на сукупності тестів."
        )

    return summary


# ============================================================
# 13. ACF / PACF
# ============================================================

def acf_pacf_analysis(data):

    print("\n" + "=" * 100)
    print("9. ACF / PACF")
    print("=" * 100)

    series = data["IT_exports"].dropna()

    max_lag = min(
        10,
        len(series) // 2 - 1
    )

    acf_values = acf(
        series,
        nlags=max_lag
    )

    pacf_values = pacf(
        series,
        nlags=max_lag
    )

    print("\nACF:")

    for lag, value in enumerate(acf_values):

        print(
            f"Lag {lag}: "
            f"{value:.6f}"
        )

    print("\nPACF:")

    for lag, value in enumerate(pacf_values):

        print(
            f"Lag {lag}: "
            f"{value:.6f}"
        )

    # ACF
    plt.figure(figsize=(10, 5))

    plt.stem(
        range(len(acf_values)),
        acf_values
    )

    plt.title(
        "ACF of IT_exports"
    )

    plt.xlabel("Lag")
    plt.ylabel("ACF")

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/ACF_IT_exports.png",
        dpi=300
    )

    plt.close()

    # PACF
    plt.figure(figsize=(10, 5))

    plt.stem(
        range(len(pacf_values)),
        pacf_values
    )

    plt.title(
        "PACF of IT_exports"
    )

    plt.xlabel("Lag")
    plt.ylabel("PACF")

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/PACF_IT_exports.png",
        dpi=300
    )

    plt.close()

    return (
        acf_values,
        pacf_values
    )


# ============================================================
# 14. ПОШУК КАНДИДАТНИХ ARDL-МОДЕЛЕЙ
# ============================================================

def ardl_lag_search(
    data,
    search_max_lag
):

    print("\n" + "=" * 100)
    print("10. ПОШУК КАНДИДАТНИХ ARDL-МОДЕЛЕЙ")
    print("=" * 100)

    y = data[
        "IT_exports"
    ]

    X = data[
        [
            "Education",
            "GERD"
        ]
    ]

    print(
        f"\nВерхня межа автоматичного перебору: "
        f"{search_max_lag}"
    )

    print(
        """
Ця межа використовується лише для технічного
перебору кандидатних моделей.

Вона НЕ визначає остаточний лаг.

Остаточну модель дослідник обирає самостійно
на основі економічної теорії, статистичних
показників та результатів діагностичних тестів.
"""
    )

    results = []

    errors = {}

    total_models = 0

    for p in range(
        1,
        search_max_lag + 1
    ):

        for q_education in range(
            0,
            search_max_lag + 1
        ):

            for q_gerd in range(
                0,
                search_max_lag + 1
            ):

                total_models += 1

                try:

                    model = ARDL(
                        y,
                        lags=p,
                        exog=X,
                        order={
                            "Education": q_education,
                            "GERD": q_gerd
                        },
                        trend="c"
                    )

                    result = model.fit()

                    results.append(
                        {
                            "p": p,
                            "q_Education": q_education,
                            "q_GERD": q_gerd,
                            "R2": result.rsquared,
                            "Adj_R2": calculate_adjusted_r2(result),
                            "LogLik": result.llf,
                            "AIC": result.aic,
                            "BIC": result.bic,
                            "HQIC": result.hqic,
                            "N_obs": result.nobs,
                            "Parameters": len(result.params)
                        }
                    )

                except Exception as e:

                    error_text = (
                        f"{type(e).__name__}: "
                        f"{str(e)}"
                    )

                    errors[error_text] = (
                        errors.get(
                            error_text,
                            0
                        ) + 1
                    )

    lag_df = pd.DataFrame(
        results
    )

    print(
        f"\nПеревірено кандидатних моделей: "
        f"{total_models}"
    )

    print(
        f"Успішно оцінено: "
        f"{len(lag_df)}"
    )

    print(
        f"Не оцінено: "
        f"{total_models - len(lag_df)}"
    )

    if lag_df.empty:

        print(
            "\nНе вдалося оцінити жодної моделі."
        )

        print(
            "\nПричини:"
        )

        for error, count in list(
            errors.items()
        )[:10]:

            print(
                f"{count} разів: {error}"
            )

        return None

    # --------------------------------------------------------
    # Виведення таблиць
    # --------------------------------------------------------

    print(
        "\nКандидатні моделі, відсортовані за AIC:"
    )

    print(
        lag_df.sort_values(
            "AIC"
        ).head(20).to_string(
            index=False
        )
    )

    print(
        "\nКандидатні моделі, відсортовані за BIC:"
    )

    print(
        lag_df.sort_values(
            "BIC"
        ).head(20).to_string(
            index=False
        )
    )

    print(
        "\nКандидатні моделі, відсортовані за HQIC:"
    )

    print(
        lag_df.sort_values(
            "HQIC"
        ).head(20).to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # Повна таблиця
    # --------------------------------------------------------

    lag_df = lag_df.sort_values(
        [
            "AIC"
        ]
    ).reset_index(
        drop=True
    )

    lag_df.to_csv(
        "ARDL_results/ARDL_candidate_models.csv",
        index=False,
        encoding="utf-8-sig"
    )

    print(
        "\nПовну таблицю кандидатних моделей "
        "збережено:"
    )

    print(
        "ARDL_results/ARDL_candidate_models.csv"
    )

    # --------------------------------------------------------
    # Інформаційні рекомендації
    # --------------------------------------------------------

    print(
        """
ВАЖЛИВО:

Моделі з мінімальними AIC, BIC або HQIC
НЕ обираються програмою як остаточна модель.

Ці критерії подаються лише як статистична
інформація для дослідника.

Остаточний вибір лагів здійснюється вручну.
"""
    )

    return lag_df


# ============================================================
# 15. РУЧНИЙ ВИБІР ЛАГІВ
# ============================================================

def manual_lag_selection():

    print("\n" + "=" * 100)
    print("11. РУЧНИЙ ВИБІР ОСТАТОЧНОЇ ARDL-МОДЕЛІ")
    print("=" * 100)

    print(
        """
Введіть лаги для остаточної моделі.

p  = лаг залежної змінної IT_exports
q1 = лаг Education
q2 = лаг GERD

Остаточний вибір робить дослідник.
"""
    )

    while True:

        try:

            p = int(
                input(
                    "\np = "
                )
            )

            q_education = int(
                input(
                    "q_Education = "
                )
            )

            q_gerd = int(
                input(
                    "q_GERD = "
                )
            )

            if p < 1:

                print(
                    "p має бути не менше 1."
                )

                continue

            if (
                q_education < 0
                or q_gerd < 0
            ):

                print(
                    "q має бути 0 або більше."
                )

                continue

            return (
                p,
                q_education,
                q_gerd
            )

        except ValueError:

            print(
                "Введіть цілі числа."
            )


# ============================================================
# 16. ОЦІНЮВАННЯ ОСТАТОЧНОЇ ARDL
# ============================================================

def fit_ardl(
    data,
    p,
    q_education,
    q_gerd
):

    print("\n" + "=" * 100)
    print("12. ОЦІНЮВАННЯ ОСТАТОЧНОЇ ARDL-МОДЕЛІ")
    print("=" * 100)

    print(
        f"\nARDL("
        f"{p}, "
        f"{q_education}, "
        f"{q_gerd}"
        f")"
    )

    y = data[
        "IT_exports"
    ]

    X = data[
        [
            "Education",
            "GERD"
        ]
    ]

    try:

        model = ARDL(
            y,
            lags=p,
            exog=X,
            order={
                "Education": q_education,
                "GERD": q_gerd
            },
            trend="c"
        )

        result = model.fit()

    except Exception as e:

        print(
            "\nПомилка оцінювання моделі:"
        )

        print(
            f"{type(e).__name__}: {e}"
        )

        return None, None

    print(
        "\nПовний результат оцінювання:"
    )

    print(
        result.summary()
    )

    # --------------------------------------------------------
    # Основні статистичні показники
    # --------------------------------------------------------

    print(
        "\nОсновні статистичні показники:"
    )

    print(
        f"R² = {result.rsquared:.6f}"
    )

    adj_r2 = calculate_adjusted_r2(result)

    print(
        f"Adjusted R² = "
        f"{adj_r2:.6f}"
    )

    print("\nP-value коефіцієнтів:")

    pvalue_table = pd.DataFrame({
        "Coefficient": result.params.index,
        "Coefficient_value": result.params.values,
        "P_value": result.pvalues.values
    })

    print(
        pvalue_table.to_string(
            index=False,
            float_format=lambda x: f"{x:.6f}"
        )
    )

    # ------------------------------------------------------------
    # Загальний F-тест значущості коефіцієнтів
    # ------------------------------------------------------------
    print("F-statistic / F-test:")

    try:
        # Для ARDLResults statsmodels не завжди надає fvalue/f_pvalue.
        # Тому використовуємо Wald-тест для всіх коефіцієнтів,
        # крім константи.
        param_names = list(result.params.index)

        test_params = [
            name for name in param_names
            if name != "const"
        ]

        if len(test_params) > 0:
            R = np.zeros((len(test_params), len(param_names)))

            for i, name in enumerate(test_params):
                j = param_names.index(name)
                R[i, j] = 1

            wald_test = result.wald_test(R)

            statistic = float(np.asarray(wald_test.statistic).squeeze())
            p_value = float(np.asarray(wald_test.pvalue).squeeze())

            print(f"Wald statistic = {statistic:.6f}")
            print(f"Wald test p-value = {p_value:.6f}")

        else:
            print("Wald-тест неможливо виконати.")

    except Exception as e:
        print(f"Wald-тест не виконано: {type(e).__name__}: {e}")

    print(
        f"Log-Likelihood = "
        f"{result.llf:.6f}"
    )

    print(
        f"AIC = "
        f"{result.aic:.6f}"
    )

    print(
        f"BIC = "
        f"{result.bic:.6f}"
    )

    print(
        f"HQIC = "
        f"{result.hqic:.6f}"
    )

    print(
        f"N = "
        f"{result.nobs}"
    )

    # --------------------------------------------------------
    # УСІ КОЕФІЦІЄНТИ ТА p-value
    # --------------------------------------------------------

    coef_table = pd.DataFrame(
        {
            "Coefficient": result.params,
            "Std_Error": result.bse,
            "t_statistic": result.tvalues,
            "p_value": result.pvalues
        }
    )

    print(
        "\nУсі коефіцієнти та p-value:"
    )

    print(
        coef_table.to_string()
    )

    # ============================================================
    # ОСТАТОЧНА ФОРМА ARDL-МОДЕЛІ
    # ============================================================

    print_final_ardl_equation(
        result,
        p=p,
        q_education=q_education,
        q_gerd=q_gerd
    )

    return model, result


# ============================================================
# 17. VIF
# ============================================================

def vif_analysis(data):

    print("\n" + "=" * 100)
    print("13. VIF")
    print("=" * 100)

    X = data[
        [
            "Education",
            "GERD"
        ]
    ].copy()

    X = X.dropna()

    vif_data = pd.DataFrame()

    vif_data["Variable"] = X.columns

    vif_data["VIF"] = [
        variance_inflation_factor(
            X.values,
            i
        )
        for i in range(
            X.shape[1]
        )
    ]

    print(
        vif_data
    )

    return vif_data


# ============================================================
# 18. BOUNDS TEST
# ============================================================

def bounds_test(model):

    print("\n" + "=" * 100)
    print("14. ARDL BOUNDS TEST")
    print("=" * 100)

    try:

        uecm = UECM.from_ardl(model)

        uecm_result = uecm.fit()

        bounds = uecm_result.bounds_test(
            case=3
        )

        print("\nПовний результат Bounds Test:")
        print(bounds)

        # ----------------------------------------------------
        # Сумісність з різними версіями statsmodels
        # ----------------------------------------------------

        if hasattr(bounds, "statistic"):
            f_stat = bounds.statistic
        elif hasattr(bounds, "stat"):
            f_stat = bounds.stat
        else:
            f_stat = np.nan

        if hasattr(bounds, "critical_values"):
            critical_values = bounds.critical_values
        elif hasattr(bounds, "crit_vals"):
            critical_values = bounds.crit_vals
        else:
            critical_values = None

        if hasattr(bounds, "pvalue"):
            p_values = bounds.pvalue
        elif hasattr(bounds, "p_values"):
            p_values = bounds.p_values
        else:
            p_values = None

        print(
            f"\nF-statistic = {float(f_stat):.6f}"
        )

        print("\nКритичні значення:")
        print(critical_values)

        print("\np-value:")
        print(p_values)

        # ----------------------------------------------------
        # Інтерпретація
        # ----------------------------------------------------

        print("\nІнтерпретація Bounds Test:")

        if p_values is not None:

            try:

                lower_p = float(p_values["lower"])
                upper_p = float(p_values["upper"])

                print(
                    f"Lower-bound p-value = {lower_p:.6g}"
                )

                print(
                    f"Upper-bound p-value = {upper_p:.6g}"
                )

            except Exception:

                print(
                    "p-value отримано у форматі:"
                )

                print(p_values)

        return (
            uecm,
            uecm_result,
            bounds
        )

    except Exception as e:

        print(
            "\nПомилка Bounds Test:"
        )

        print(
            f"{type(e).__name__}: {e}"
        )

        return (
            None,
            None,
            None
        )

# ============================================================
# 19. LONG-RUN ANALYSIS
# ============================================================

def long_run_analysis(
    uecm_result
):

    print("\n" + "=" * 100)
    print("15. ДОВГОСТРОКОВІ КОЕФІЦІЄНТИ")
    print("=" * 100)

    if uecm_result is None:

        print(
            "UECM результат відсутній."
        )

        return None

    try:

        # ----------------------------------------------------
        # 1. Нормалізований коінтеграційний зв'язок
        # ----------------------------------------------------

        print(
            "\nНормалізований коінтеграційний зв'язок:"
        )

        print(
            uecm_result.ci_summary()
        )

        # ----------------------------------------------------
        # 2. Отримання коефіцієнтів UECM
        # ----------------------------------------------------

        params = uecm_result.params
        pvalues = uecm_result.pvalues
        bse = uecm_result.bse
        tvalues = uecm_result.tvalues

        # ----------------------------------------------------
        # 3. Коефіцієнт при лагованій залежній змінній
        # ----------------------------------------------------

        y_lag_name = "IT_exports.L1"

        if y_lag_name not in params.index:

            print(
                "\nПомилка: не знайдено коефіцієнт "
                "IT_exports.L1."
            )

            return None

        phi = params[y_lag_name]

        denominator = 1 - phi

        if abs(denominator) < 1e-10:

            print(
                "\nНеможливо розрахувати "
                "довгострокові коефіцієнти:"
            )

            print(
                "1 - коефіцієнт IT_exports.L1 "
                "занадто близький до нуля."
            )

            return None

        # ----------------------------------------------------
        # 4. Long-run constant
        # ----------------------------------------------------

        if "const" in params.index:

            lr_const = (
                params["const"] /
                denominator
            )

        else:

            lr_const = np.nan

        # ----------------------------------------------------
        # 5. Long-run Education
        # ----------------------------------------------------

        education_coefficients = [
            name
            for name in params.index
            if name.startswith("Education.")
        ]

        education_sum = 0.0

        for name in education_coefficients:

            education_sum += params[name]

        lr_education = (
            education_sum /
            denominator
        )

        # ----------------------------------------------------
        # 6. Long-run GERD
        # ----------------------------------------------------

        gerd_coefficients = [
            name
            for name in params.index
            if name.startswith("GERD.")
        ]

        gerd_sum = 0.0

        for name in gerd_coefficients:

            gerd_sum += params[name]

        lr_gerd = (
            gerd_sum /
            denominator
        )

        # ----------------------------------------------------
        # 7. P-values
        # ----------------------------------------------------
        #
        # Для довгострокових коефіцієнтів використовуємо
        # p-value з нормалізованого коінтеграційного зв'язку.
        #
        # Зміна знака коефіцієнта не змінює двосторонній
        # p-value.
        # ----------------------------------------------------

        try:

            ci_pvalues = pd.Series(
                uecm_result.ci_pvalues
            )

        except Exception:

            ci_pvalues = pd.Series(
                dtype=float
            )

        education_pvalue = np.nan
        gerd_pvalue = np.nan

        for name in ci_pvalues.index:

            if "Education" in str(name):

                education_pvalue = (
                    ci_pvalues[name]
                )

            if "GERD" in str(name):

                gerd_pvalue = (
                    ci_pvalues[name]
                )

        # ----------------------------------------------------
        # 8. Формування таблиці
        # ----------------------------------------------------

        table = pd.DataFrame(
            {
                "Variable": [
                    "const",
                    "Education",
                    "GERD"
                ],

                "Long_run_coefficient": [
                    lr_const,
                    lr_education,
                    lr_gerd
                ],

                "p_value": [
                    np.nan,
                    education_pvalue,
                    gerd_pvalue
                ]
            }
        )

        # ----------------------------------------------------
        # 9. Виведення
        # ----------------------------------------------------

        print(
            "\nЕкономічні довгострокові коефіцієнти:"
        )

        print(
            table.to_string(
                index=False
            )
        )

        # ----------------------------------------------------
        # 10. Long-run equation
        # ----------------------------------------------------

        print(
            "\n" + "-" * 100
        )

        print(
            "ДОВГОСТРОКОВЕ РІВНЯННЯ"
        )

        print(
            "-" * 100
        )

        print(
            "\nIT_exports = "
            f"{lr_const:.6f}"
            f" {lr_education:+.6f} * Education"
            f" {lr_gerd:+.6f} * GERD"
        )

        # ----------------------------------------------------
        # 11. Економічна інтерпретація
        # ----------------------------------------------------

        print(
            "\n" + "-" * 100
        )

        print(
            "ЕКОНОМІЧНА ІНТЕРПРЕТАЦІЯ"
        )

        print(
            "-" * 100
        )

        print(
            f"\nEducation:"
        )

        print(
            f"Long-run coefficient = "
            f"{lr_education:.6f}"
        )

        if not np.isnan(education_pvalue):

            print(
                f"p-value = "
                f"{education_pvalue:.6f}"
            )

            if education_pvalue < 0.05:

                print(
                    "Коефіцієнт статистично значущий "
                    "на рівні 5%."
                )

            else:

                print(
                    "Коефіцієнт статистично незначущий "
                    "на рівні 5%."
                )

        print(
            "\nGERD:"
        )

        print(
            f"Long-run coefficient = "
            f"{lr_gerd:.6f}"
        )

        if not np.isnan(gerd_pvalue):

            print(
                f"p-value = "
                f"{gerd_pvalue:.6f}"
            )

            if gerd_pvalue < 0.05:

                print(
                    "Коефіцієнт статистично значущий "
                    "на рівні 5%."
                )

            else:

                print(
                    "Коефіцієнт статистично незначущий "
                    "на рівні 5%."
                )

        # ----------------------------------------------------
        # 12. Важливе пояснення щодо знаків
        # ----------------------------------------------------

        print(
            "\nПримітка:"
        )

        print(
            "Наведені вище коефіцієнти є економічними "
            "довгостроковими ефектами, отриманими шляхом "
            "нормалізації рівняння відносно IT_exports."
        )

        print(
            "Вони можуть мати протилежні знаки порівняно "
            "з коефіцієнтами нормалізованого "
            "коінтеграційного вектора."
        )

        return table

    except Exception as e:

        print(
            "\nПомилка розрахунку "
            "довгострокових коефіцієнтів:"
        )

        print(
            f"{type(e).__name__}: {e}"
        )

        return None

# ============================================================
# 20. ECM
# ============================================================

def ecm_analysis(
    uecm_result
):

    print("\n" + "=" * 100)
    print("16. ECM")
    print("=" * 100)

    if uecm_result is None:

        print(
            "UECM результат відсутній."
        )

        return None

    try:

        print(
            "\nПовний результат UECM / ECM:"
        )

        print(
            uecm_result.summary()
        )

        # ----------------------------------------------------
        # Таблиця коефіцієнтів
        # ----------------------------------------------------

        ecm_table = pd.DataFrame(
            {
                "Coefficient": uecm_result.params,
                "Std_Error": uecm_result.bse,
                "t_statistic": uecm_result.tvalues,
                "p_value": uecm_result.pvalues
            }
        )

        print(
            "\nКоефіцієнти ECM:"
        )

        print(
            ecm_table.to_string()
        )

        # ----------------------------------------------------
        # Error Correction Coefficient
        # ----------------------------------------------------

        level_params = [
            name
            for name in uecm_result.params.index
            if (
                name.startswith("IT_exports.L1")
            )
        ]

        if level_params:

            ecm_coef_name = level_params[0]

            ecm_coef = (
                uecm_result.params[
                    ecm_coef_name
                ]
            )

            ecm_pvalue = (
                uecm_result.pvalues[
                    ecm_coef_name
                ]
            )

            print(
                "\nКоефіцієнт корекції помилки:"
            )

            print(
                f"{ecm_coef_name} = "
                f"{ecm_coef:.6f}"
            )

            print(
                f"p-value = "
                f"{ecm_pvalue:.6f}"
            )

            if ecm_coef < 0:

                print(
                    "\nКоефіцієнт корекції помилки "
                    "має від'ємний знак."
                )

                if ecm_pvalue < 0.05:

                    print(
                        "Він статистично значущий "
                        "на рівні 5%, що є узгодженим "
                        "із механізмом повернення системи "
                        "до довгострокової рівноваги."
                    )

                else:

                    print(
                        "Він не є статистично значущим "
                        "на рівні 5%."
                    )

            else:

                print(
                    "\nКоефіцієнт корекції помилки "
                    "не має від'ємного знака."
                )

        return ecm_table

    except Exception as e:

        print(
            "\nПомилка ECM:"
        )

        print(
            f"{type(e).__name__}: {e}"
        )

        return None

# ============================================================
# 21. ДІАГНОСТИКА ЗАЛИШКІВ
# ============================================================

def residual_diagnostics(
    model,
    result
):

    print("\n" + "=" * 100)
    print("17. ДІАГНОСТИКА ЗАЛИШКІВ")
    print("=" * 100)

    resid = pd.Series(
        result.resid
    ).dropna()

    diagnostics = {}

    # --------------------------------------------------------
    # Jarque-Bera
    # --------------------------------------------------------

    jb_stat, jb_p, skew, kurtosis = jarque_bera(
        resid
    )

    diagnostics["JB_stat"] = jb_stat
    diagnostics["JB_p_value"] = jb_p
    diagnostics["Skewness"] = skew
    diagnostics["Kurtosis"] = kurtosis

    print(
        "\nJarque-Bera:"
    )

    print(
        f"Statistic = {jb_stat:.6f}"
    )

    print(
        f"p-value = {jb_p:.6f}"
    )

    # --------------------------------------------------------
    # Breusch-Godfrey
    # --------------------------------------------------------

    for lag in [1, 2, 3]:

        try:

            bg_stat, bg_p, bg_aux = (
                breusch_godfrey_ardl(
                    result,
                    lag
                )
            )

            diagnostics[
                f"BG_{lag}_stat"
            ] = bg_stat

            diagnostics[
                f"BG_{lag}_p_value"
            ] = bg_p

            print(
                f"\nBreusch-Godfrey "
                f"(lag {lag}):"
            )

            print(
                f"LM statistic = "
                f"{bg_stat:.6f}"
            )

            print(
                f"p-value = "
                f"{bg_p:.6f}"
            )

        except Exception as e:

            print(
                f"\nBG lag {lag}: "
                f"помилка "
                f"{type(e).__name__}: {e}"
            )

    # --------------------------------------------------------
    # Breusch-Pagan
    # --------------------------------------------------------

    try:

        bp = het_breuschpagan(
            resid,
            result.model._x
        )

        diagnostics[
            "BP_stat"
        ] = bp[0]

        diagnostics[
            "BP_p_value"
        ] = bp[1]

        print(
            "\nBreusch-Pagan:"
        )

        print(
            f"Statistic = {bp[0]:.6f}"
        )

        print(
            f"p-value = {bp[1]:.6f}"
        )

    except Exception as e:

        print(
            "\nBreusch-Pagan:"
        )

        print(
            f"Не вдалося виконати: {e}"
        )

    # --------------------------------------------------------
    # White
    # --------------------------------------------------------

    try:

        white = het_white(
            resid,
            result.model._x
        )

        diagnostics[
            "White_stat"
        ] = white[0]

        diagnostics[
            "White_p_value"
        ] = white[1]

        print(
            "\nWhite test:"
        )

        print(
            f"Statistic = {white[0]:.6f}"
        )

        print(
            f"p-value = {white[1]:.6f}"
        )

    except Exception as e:

        print(
            "\nWhite test:"
        )

        print(
            f"Не вдалося виконати: {e}"
        )

    # --------------------------------------------------------
    # Графік залишків
    # --------------------------------------------------------

    plt.figure(figsize=(11, 5))

    plt.plot(
        resid.index,
        resid.values
    )

    plt.axhline(
        0,
        linestyle="--"
    )

    plt.title(
        "ARDL residuals"
    )

    plt.xlabel("Observation")
    plt.ylabel("Residual")

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/residuals.png",
        dpi=300
    )

    plt.close()

    # --------------------------------------------------------
    # Гістограма
    # --------------------------------------------------------

    plt.figure(figsize=(9, 5))

    plt.hist(
        resid,
        bins=10
    )

    plt.title(
        "Distribution of ARDL residuals"
    )

    plt.xlabel("Residual")
    plt.ylabel("Frequency")

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/residual_histogram.png",
        dpi=300
    )

    plt.close()

    return diagnostics


# ============================================================
# 22. RESET
# ============================================================

def ramsey_reset_test(
    result
):

    print("\n" + "=" * 100)
    print("18. RAMSEY RESET")
    print("=" * 100)

    try:

        reset_result = linear_reset(
            result,
            power=2,
            use_f=True
        )

        print(
            reset_result
        )

        print(
            f"\nF-statistic = "
            f"{reset_result.fvalue}"
        )

        print(
            f"p-value = "
            f"{reset_result.pvalue}"
        )

        return reset_result

    except Exception as e:

        print(
            "\nНе вдалося виконати RESET:"
        )

        print(
            f"{type(e).__name__}: {e}"
        )

        return None


# ============================================================
# 23. CUSUM / CUSUMSQ
# ============================================================

def cusum_tests(
    result
):

    print("\n" + "=" * 100)
    print("19. CUSUM / CUSUMSQ")
    print("=" * 100)

    resid = pd.Series(
        result.resid
    ).dropna()

    std_resid = (
        resid - resid.mean()
    ) / resid.std()

    cusum = np.cumsum(
        std_resid
    )

    cusumsq = np.cumsum(
        std_resid ** 2
    )

    cusumsq = (
        cusumsq /
        cusumsq.iloc[-1]
    )

    # --------------------------------------------------------
    # CUSUM
    # --------------------------------------------------------

    plt.figure(figsize=(11, 5))

    plt.plot(
        cusum
    )

    plt.axhline(
        0,
        linestyle="--"
    )

    plt.title(
        "CUSUM of standardized residuals"
    )

    plt.xlabel("Observation")
    plt.ylabel("CUSUM")

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/CUSUM.png",
        dpi=300
    )

    plt.close()

    # --------------------------------------------------------
    # CUSUMSQ
    # --------------------------------------------------------

    plt.figure(figsize=(11, 5))

    plt.plot(
        cusumsq
    )

    plt.axhline(
        0.5,
        linestyle="--"
    )

    plt.axhline(
        1,
        linestyle="--"
    )

    plt.title(
        "CUSUMSQ of standardized residuals"
    )

    plt.xlabel("Observation")
    plt.ylabel("CUSUMSQ")

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/CUSUMSQ.png",
        dpi=300
    )

    plt.close()

    print(
        """
Графіки CUSUM та CUSUMSQ збережено.

УВАГА: ці графіки є діагностикою стабільності
залишків. Вони не замінюють формальну процедуру
CUSUM/CUSUMSQ з критичними межами Pesaran et al.
"""
    )


# ============================================================
# 24. ФАКТИЧНІ ТА ОЦІНЕНІ ЗНАЧЕННЯ
# ============================================================

def actual_vs_fitted(
    data,
    result
):

    print("\n" + "=" * 100)
    print("20. ФАКТИЧНІ ТА ОЦІНЕНІ ЗНАЧЕННЯ")
    print("=" * 100)

    fitted = result.fittedvalues

    comparison = pd.DataFrame(
        {
            "Year": data.loc[
                fitted.index,
                "Year"
            ],
            "Actual_IT_exports":
                data.loc[
                    fitted.index,
                    "IT_exports"
                ],
            "Fitted_IT_exports":
                fitted
        }
    )

    print(
        comparison.to_string(
            index=False
        )
    )

    plt.figure(figsize=(11, 5))

    plt.plot(
        comparison["Year"],
        comparison["Actual_IT_exports"],
        marker="o",
        label="Actual"
    )

    plt.plot(
        comparison["Year"],
        comparison["Fitted_IT_exports"],
        marker="o",
        label="Fitted"
    )

    plt.xlabel("Year")
    plt.ylabel("IT_exports")

    plt.title(
        "Actual vs Fitted values"
    )

    plt.legend()

    plt.grid(
        True,
        alpha=0.3
    )

    plt.tight_layout()

    plt.savefig(
        "ARDL_results/actual_vs_fitted.png",
        dpi=300
    )

    plt.close()

    return comparison


# ============================================================
# 25. EXCEL REPORT
# ============================================================

def create_excel_report(
    data,
    descriptive,
    pearson_matrix,
    spearman_matrix,
    correlation_tests,
    level_results,
    diff_results,
    second_diff_results,
    second_diff_summary,
    integration_summary_df,
    lag_df,
    coef_table,
    vif_data,
    long_run_table,
    diagnostics,
    comparison
):

    print("\n" + "=" * 100)
    print("21. ФОРМУВАННЯ EXCEL-ЗВІТУ")
    print("=" * 100)

    filename = (
        "ARDL_results/"
        "ARDL_analysis_report.xlsx"
    )

    with pd.ExcelWriter(
        filename,
        engine="openpyxl"
    ) as writer:

        data.to_excel(
            writer,
            sheet_name="Data",
            index=False
        )

        descriptive.to_excel(
            writer,
            sheet_name="Descriptive"
        )

        pearson_matrix.to_excel(
            writer,
            sheet_name="Pearson"
        )

        spearman_matrix.to_excel(
            writer,
            sheet_name="Spearman"
        )

        correlation_tests.to_excel(
            writer,
            sheet_name="Correlation_tests",
            index=False
        )

        level_results.to_excel(
            writer,
            sheet_name="Unit_root_levels",
            index=False
        )

        diff_results.to_excel(
            writer,
            sheet_name="Unit_root_1diff",
            index=False
        )

        second_diff_results.to_excel(
            writer,
            sheet_name="Unit_root_2diff",
            index=False
        )

        second_diff_summary.to_excel(
            writer,
            sheet_name="I2_summary",
            index=False
        )

        integration_summary_df.to_excel(
            writer,
            sheet_name="Integration_summary",
            index=False
        )

        if lag_df is not None:

            lag_df.to_excel(
                writer,
                sheet_name="ARDL_candidates",
                index=False
            )

        if coef_table is not None:

            coef_table.to_excel(
                writer,
                sheet_name="Final_coefficients"
            )

        if vif_data is not None:

            vif_data.to_excel(
                writer,
                sheet_name="VIF",
                index=False
            )

        if long_run_table is not None:

            long_run_table.to_excel(
                writer,
                sheet_name="Long_run"
            )

        if diagnostics:

            pd.DataFrame(
                [diagnostics]
            ).to_excel(
                writer,
                sheet_name="Diagnostics",
                index=False
            )

        if comparison is not None:

            comparison.to_excel(
                writer,
                sheet_name="Actual_vs_fitted",
                index=False
            )

    # --------------------------------------------------------
    # Форматування
    # --------------------------------------------------------

    workbook = load_workbook(
        filename
    )

    for worksheet in workbook.worksheets:

        worksheet.freeze_panes = "A2"

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = get_column_letter(
                column_cells[0].column
            )

            for cell in column_cells:

                try:

                    length = len(
                        str(cell.value)
                    )

                    if length > max_length:
                        max_length = length

                except Exception:
                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                35
            )

    workbook.save(
        filename
    )

    print(
        f"\nExcel-звіт збережено:"
    )

    print(
        filename
    )


# ============================================================
# 26. MAIN
# ============================================================

def main():

    print("\n")
    print("=" * 100)
    print("ARDL ECONOMETRIC ANALYSIS")
    print("Період: 2000–2024")
    print("=" * 100)

    os.makedirs(
        "ARDL_results",
        exist_ok=True
    )

    # --------------------------------------------------------
    # 1. Дані
    # --------------------------------------------------------

    data = load_excel_data()

    # --------------------------------------------------------
    # 2. Описова статистика
    # --------------------------------------------------------

    descriptive = descriptive_statistics(
        data
    )

    # --------------------------------------------------------
    # 3. Графіки
    # --------------------------------------------------------

    plot_time_series(
        data
    )

    # --------------------------------------------------------
    # 4. Кореляція
    # --------------------------------------------------------

    (
        pearson_matrix,
        spearman_matrix,
        correlation_tests
    ) = correlation_analysis(
        data
    )

    # --------------------------------------------------------
    # 5. Рівні
    # --------------------------------------------------------

    level_results = unit_root_tests(
        data
    )

    # --------------------------------------------------------
    # 6. Перші різниці
    # --------------------------------------------------------

    diff_results = first_difference_tests(
        data
    )

    # --------------------------------------------------------
    # 7. Другі різниці / I(2)
    # --------------------------------------------------------

    (
        second_diff_results,
        second_diff_summary
    ) = second_difference_tests(
        data
    )

    # --------------------------------------------------------
    # 8. Порядок інтеграції
    # --------------------------------------------------------

    integration_summary_df = integration_summary(
        level_results,
        diff_results,
        second_diff_summary
    )

    # --------------------------------------------------------
    # 9. ACF/PACF
    # --------------------------------------------------------

    acf_values, pacf_values = acf_pacf_analysis(
        data
    )

    # --------------------------------------------------------
    # 10. Межа автоматичного перебору
    # --------------------------------------------------------

    print("\n" + "=" * 100)
    print("ВИБІР МЕЖІ ТЕХНІЧНОГО ПЕРЕБОРУ ЛАГІВ")
    print("=" * 100)

    print(
        """
Програма НЕ встановлює економетричний MAX_LAG
як остаточне обмеження моделі.

Тут потрібно лише задати верхню межу для
автоматичного перебору кандидатних моделей.

Це технічний параметр пошуку.

Остаточний p, q_Education та q_GERD
ВИ ОБИРАЄТЕ САМОСТІЙНО.
"""
    )

    while True:

        try:

            search_max_lag = int(
                input(
                    "\nВерхня межа технічного перебору лагів: "
                )
            )

            if search_max_lag < 1:

                print(
                    "Введіть число не менше 1."
                )

                continue

            break

        except ValueError:

            print(
                "Введіть ціле число."
            )

    # --------------------------------------------------------
    # 11. Пошук
    # --------------------------------------------------------

    lag_df = ardl_lag_search(
        data,
        search_max_lag
    )

    if lag_df is None:

        print(
            "\nПрограма завершена через "
            "відсутність оцінених моделей."
        )

        return

    # --------------------------------------------------------
    # 12. Ручний вибір
    # --------------------------------------------------------

    (
        p,
        q_education,
        q_gerd
    ) = manual_lag_selection()

    # --------------------------------------------------------
    # 13. Остаточна модель
    # --------------------------------------------------------

    model, result = fit_ardl(
        data,
        p,
        q_education,
        q_gerd
    )

    if model is None:

        print(
            "\nОстаточна модель не оцінена."
        )

        return

    coef_table = pd.DataFrame(
        {
            "Coefficient": result.params,
            "Std_Error": result.bse,
            "t_statistic": result.tvalues,
            "p_value": result.pvalues
        }
    )

    # --------------------------------------------------------
    # 14. VIF
    # --------------------------------------------------------

    vif_data = vif_analysis(
        data
    )

    # --------------------------------------------------------
    # 15. Bounds
    # --------------------------------------------------------

    (
        uecm,
        uecm_result,
        bounds
    ) = bounds_test(
        model
    )

    # --------------------------------------------------------
    # 16. Long-run
    # --------------------------------------------------------

    long_run_table = long_run_analysis(
        uecm_result
    )

    # --------------------------------------------------------
    # 17. ECM
    # --------------------------------------------------------

    ecm_analysis(
        uecm_result
    )

    # --------------------------------------------------------
    # 18. Diagnostics
    # --------------------------------------------------------

    diagnostics = residual_diagnostics(
        model,
        result
    )

    # --------------------------------------------------------
    # 19. RESET
    # --------------------------------------------------------

    ramsey_reset_test(
        result
    )

    # --------------------------------------------------------
    # 20. CUSUM
    # --------------------------------------------------------

    cusum_tests(
        result
    )

    # --------------------------------------------------------
    # 21. Actual vs fitted
    # --------------------------------------------------------

    comparison = actual_vs_fitted(
        data,
        result
    )

    # --------------------------------------------------------
    # 22. Excel
    # --------------------------------------------------------

    create_excel_report(
        data=data,
        descriptive=descriptive,
        pearson_matrix=pearson_matrix,
        spearman_matrix=spearman_matrix,
        correlation_tests=correlation_tests,
        level_results=level_results,
        diff_results=diff_results,
        second_diff_results=second_diff_results,
        second_diff_summary=second_diff_summary,
        integration_summary_df=integration_summary_df,
        lag_df=lag_df,
        coef_table=coef_table,
        vif_data=vif_data,
        long_run_table=long_run_table,
        diagnostics=diagnostics,
        comparison=comparison
    )

    print("\n" + "=" * 100)
    print("АНАЛІЗ ЗАВЕРШЕНО")
    print("=" * 100)

    print(
        "\nУсі результати збережені у папці:"
    )

    print(
        "ARDL_results/"
    )


# ============================================================
# ЗАПУСК
# ============================================================

if __name__ == "__main__":
    main()