# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::WorkbookProperties do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/workbook_properties/workbook_properties_date1904.xlsx')

  it 'Has date1904' do
    expect(xlsx.workbook_properties.date1904).to be_truthy
  end
end
