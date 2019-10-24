# frozen_string_literal: true

require 'spec_helper'

describe 'SharedStrings' do
  it 'custom_shared_strings_name.xlsx' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/shared_strings/custom_shared_strings_name.xlsx')
    expect(xlsx.shared_strings_table.count).to eq(1)
  end
end
