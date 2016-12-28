require 'spec_helper'

describe 'show_row_column_heading' do
  it 'show_headers.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/sheet_view/row_column_headings/show_headers.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.show_row_column_headers).to be_truthy
  end

  it 'hide_headers.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/sheet_view/row_column_headings/hide_headers.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.show_row_column_headers).to be_falsey
  end
end
