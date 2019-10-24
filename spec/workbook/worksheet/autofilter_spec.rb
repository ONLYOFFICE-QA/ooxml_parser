# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'autofilter.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/autofilter/autofilter.xlsx')
    expect(xlsx.worksheets.first.autofilter.ref.first.column).to eq('A')
    expect(xlsx.worksheets.first.autofilter.ref.first.row).to eq(1)
  end
end
