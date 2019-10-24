# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Pane do
  it 'show_gridlines.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/sheet_view/gridlines/show_gridlines.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.show_gridlines).to be_truthy
  end

  it 'hide_gridlines.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/sheet_view/gridlines/hide_gridlines.xlsx')
    expect(xlsx.worksheets.first.sheet_views.first.show_gridlines).to be_falsey
  end
end
