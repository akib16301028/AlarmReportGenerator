def style_dataframe(df, duration_cols, is_dark_mode):
    # Replace 0 with empty strings in duration columns
    df_style = df.copy()
    df_style[duration_cols] = df_style[duration_cols].replace(0, "")

    # Define background colors based on theme
    if is_dark_mode:
        duration_bg = '#E9D2D2'  # Very light gray suitable for dark mode
        other_bg = '#ADD8E6'      # Light blue suitable for dark mode
        total_bg = '#B0B0B0'      # Very light grey for total row in dark mode
        font_color = 'white'
    else:
        duration_bg = '#EBADAD'  # Very light gray
        other_bg = '#ADD8E6'      # Very light blue
        total_bg = '#D3D3D3'      # Very light grey for total row in light mode
        font_color = 'black'

    # Create Styler object
    styler = df_style.style

    # Apply background color to duration columns
    styler = styler.applymap(
        lambda x: f'background-color: {duration_bg}; color: {font_color}' if x != "" else '',
        subset=duration_cols
    )

    # Apply background color to other columns (excluding Cluster and Zone)
    non_duration_cols = [col for col in df_style.columns if col not in ['Cluster', 'Zone'] + duration_cols]
    styler = styler.applymap(
        lambda x: f'background-color: {other_bg}; color: {font_color}' if pd.notna(x) and x != "" else '',
        subset=non_duration_cols
    )

    # Apply special background color for the total row
    total_row_index = df[df['Cluster'] == 'Total'].index
    styler = styler.apply(
        lambda row: ['background-color: {}'.format(total_bg) if row.name in total_row_index else '' for _ in row],
        axis=1
    )

    # Hide borders for a cleaner look
    styler.set_table_styles(
        [
            {
                'selector': 'th',
                'props': [('border', '1px solid black')]
            },
            {
                'selector': 'td',
                'props': [('border', '1px solid black')]
            }
        ]
    )

    return styler
