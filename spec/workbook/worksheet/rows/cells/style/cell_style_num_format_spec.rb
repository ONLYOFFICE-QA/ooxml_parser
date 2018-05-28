require 'spec_helper'

describe 'cell_style_num_format' do
  it 'cell_style_num_format_enabled' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_style_num_format/cell_style_num_format_enabled.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.apply_number_format).to be_truthy
  end

  it 'cell_style_num_format_disabled' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_style_num_format/cell_style_num_format_disabled.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.apply_number_format).to be_falsey
  end

  it 'cell_style_num_format_value' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_style_num_format/cell_style_num_format_value.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.numerical_format).to eq('h:mm:ss')
  end
end
