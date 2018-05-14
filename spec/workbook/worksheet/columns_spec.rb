require 'spec_helper'

describe 'My behaviour' do
  it 'columns_width_custom.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/columns/columns_width_custom.xlsx')
    expect(xlsx.worksheets.first.columns.first.custom_width).to be_truthy
  end

  it 'column_best_fit.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/columns/column_best_fit.xlsx')
    expect(xlsx.worksheets.first.columns.first.best_fit).to be_truthy
  end

  it 'column_hidden.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/columns/column_hidden.xlsx')
    expect(xlsx.worksheets.first.columns.first.hidden).to be_truthy
  end
end
