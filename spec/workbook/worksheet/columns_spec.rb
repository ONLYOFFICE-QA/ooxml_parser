require 'spec_helper'

describe 'My behaviour' do
  it 'columns_width_custom.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/columns/columns_width_custom.xlsx')
    expect(xlsx.worksheets.first.columns.first.custom_width).to be_truthy
  end
end
