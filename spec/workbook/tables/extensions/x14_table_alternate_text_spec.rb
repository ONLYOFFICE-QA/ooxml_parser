# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::X14Table do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/tables/' \
                              'extensions/x14_table/' \
                              'x14_table_alternative_text.xlsx')
  end

  it 'alt text' do
    expect(xlsx.worksheets.first.table_parts.first.extension_list[0].table.alt_text).to eq('ONLYOFFICE Title Support')
  end

  it 'alt text summary' do
    expect(xlsx.worksheets.first.table_parts.first.extension_list[0].table.alt_text_summary).to eq('ONLYOFFICE Title Description')
  end
end
