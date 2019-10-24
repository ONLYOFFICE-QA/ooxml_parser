# frozen_string_literal: true

require 'spec_helper'

describe 'Autofilter#filter_column' do
  it 'Filter column show button is enabled by default' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/autofilter/filter_column/show_filter_button_enabled.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.autofilter.filter_column.show_button).to be_truthy
  end

  it 'Filter column show button is disabled' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/autofilter/filter_column/show_filter_button_disabled.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.autofilter.filter_column.show_button).to be_falsey
  end
end
