# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'theme_without_name' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/theme/theme_without_name.xlsx')
    expect(xlsx.theme.name).to be_empty
  end
end
