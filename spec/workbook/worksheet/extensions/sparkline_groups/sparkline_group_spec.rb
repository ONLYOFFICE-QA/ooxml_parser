require 'spec_helper'

describe 'Sparkline group' do
  it 'sparkline_group_type.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_group_type.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].type).to eq(:column)
  end

  it 'sparkline_line_weight.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_line_weight.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].line_weight).to eq(OoxmlParser::OoxmlSize.new(0.5, :point))
  end

  it 'sparkline_show_points.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_show_points.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].high_point).to be_truthy
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].low_point).to be_truthy
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].negative_point).to be_truthy
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].first_point).to be_truthy
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].last_point).to be_truthy
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].markers).to be_truthy
  end

  it 'sparkline_points_colors.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_points_colors.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_series).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_high).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_low).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_negative).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_first).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_last).to be_a(OoxmlParser::OoxmlColor)
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].color_markers).to be_a(OoxmlParser::OoxmlColor)
  end

  it 'sparkline_display_empty_cells.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_display_empty_cells.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].display_empty_cells_as).to eq(:zero)
  end
end
