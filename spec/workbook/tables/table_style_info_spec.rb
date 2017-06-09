require 'spec_helper'

describe 'Table#table_style_info' do
  it 'Table Style info name' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/table_style_info/table_style_info_name.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.table_style_info.name).to eq('TableStyleMedium28')
  end

  it 'Table Style info columns data' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/table_style_info/table_style_info_column_stripes.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.table_style_info.show_column_stripes).to be_falsey
    expect(xlsx.worksheets.first.table_parts.first.table_style_info.show_first_column).to be_falsey
    expect(xlsx.worksheets.first.table_parts.first.table_style_info.show_last_column).to be_falsey
    expect(xlsx.worksheets.first.table_parts.first.table_style_info.show_row_stripes).to be_truthy
  end
end
