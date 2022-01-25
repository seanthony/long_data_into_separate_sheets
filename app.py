import pandas as pd


def open_treated_data(filename):
    df = pd.read_excel(filename)
    df = df[['Assoc_FY', 'Metric', 'Metric_Type', 'Region', 'Abscissa', 'Value']]

    df = df.replace("Org-Wide", "03 Org-Wide")

    return df


def sort_and_prep(df):
    df = df.sort_values(by=['Metric', 'Metric_Type', 'Region', 'Abscissa'])
    return df


def extract_fields(df):
    array_metrics = df['Metric'].unique()
    all_metrics = list(array_metrics)
    metrics = [metric for metric in all_metrics if ' KPI' not in metric]

    regions = pd.DataFrame(df["Region"].unique(), columns=['Region']) 
    front_matter = pd.DataFrame([['00 Fiscal Year'], ['01 Abscissa'], ['02 Region']], columns=['Region'])
    regions = pd.concat([front_matter, regions])
    regions = regions.reset_index()
    regions = regions[['Region']]
    
    return metrics, regions


def create_dataframes(df, metrics, regions):
    d = {}
    for metric in metrics:
        df_cut = df[df['Metric'] == metric]
        df_values = df_cut[['Region', 'Metric', 'Metric_Type', 'Value', 'Abscissa', 'Assoc_FY']]

        # initialize empty list of dfs
        l = []

        # get unique values of lists
        assoc_fys = df_values['Assoc_FY'].unique()
        for assoc_fy in assoc_fys:
            df_fy = df_values[df_values['Assoc_FY'] == assoc_fy]
            abscissas = df_fy['Abscissa'].unique()
            metric_type = list(df_fy['Metric_Type'])[0]
            for abscissa in abscissas:
                _df = df_fy[df_fy['Abscissa'] == abscissa]
                region_replacement = f'{assoc_fy} {metric_type}\n({abscissa})'
                front_matter = pd.DataFrame([['00 Fiscal Year', assoc_fy], ['01 Abscissa', abscissa], ['02 Region', region_replacement]], columns=['Region', 'Value'])
                _df = _df[['Region', 'Value']]
                _df = pd.concat([front_matter, _df])
                metric_new_name = region_replacement.replace('\n', ' ')
                _df = _df.rename(columns={'Value': metric_new_name})
                _df = _df.reset_index()
                _df = _df[['Region', metric_new_name]]
                l.append(_df.copy())

        # merge dataframes
        merged_dfs = pd.merge(regions, l[0], how='left', on='Region')
        for _df in l[1:]:
            merged_dfs = pd.merge(merged_dfs, _df, how='left', on='Region')

        # add to dictionary of dataframes
        cleaned_metric = ''.join([ch for ch in metric if ch.isalpha() or ch == ' '])
        while '  ' in cleaned_metric:
            cleaned_metric = cleaned_metric.replace('  ', ' ')
        cleaned_metric = cleaned_metric[:30]
        d[cleaned_metric] = merged_dfs.copy()

    return d


def write_output_file(d, filename):
    writer = pd.ExcelWriter(filename, engine='openpyxl')

    for sheetname, _df in d.items():
        print('\tsaving...', sheetname)
        _df.to_excel(writer, sheet_name=sheetname, index=False)

    writer.save()


def main():
    print('opening datafile'.center(60, '.'))
    df = open_treated_data('./TreatedData.xlsx')
    print('sort dataframe'.center(60, '.'))
    df = sort_and_prep(df)
    print('getting unique identifiers'.center(60, '.'))
    metrics, regions = extract_fields(df)
    print('creating dataframes'.center(60, '.'))
    d = create_dataframes(df, metrics, regions)
    print('saving excel file'.center(60, '.'))
    write_output_file(d, './output.xlsx')


if __name__ == '__main__':
    main()