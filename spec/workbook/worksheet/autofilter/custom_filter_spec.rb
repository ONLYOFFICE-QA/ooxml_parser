# frozen_string_literal: true

require 'spec_helper'

describe 'Custom Filter' do
  it 'custom_filter.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/autofilter/custom_filter/custom_filter.xlsx')
    expect(xlsx.worksheets.first.autofilter.filter_column.custom_filters[0].operator).to eq(:greater_than)
    expect(xlsx.worksheets.first.autofilter.filter_column.custom_filters[0].value).to eq(1)
  end
end
