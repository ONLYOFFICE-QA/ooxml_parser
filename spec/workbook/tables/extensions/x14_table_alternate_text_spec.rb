# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::X14Table do
  it 'alt text' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/extensions/x14_table/x14_table_alternative_text.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.extension_list[0].table.alt_text).to eq('ONLYOFFICE Title Support')
  end

  it 'alt text' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/extensions/x14_table/x14_table_alternative_text.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.extension_list[0].table.alt_text_summary).to eq('ONLYOFFICE Title Description')
  end
end
