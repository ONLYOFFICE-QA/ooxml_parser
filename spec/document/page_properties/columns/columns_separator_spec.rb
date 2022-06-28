# frozen_string_literal: true

require 'spec_helper'

describe 'Columns' do
  it 'separator is not present' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/' \
                                              'columns/columns_separator/' \
                                              'no_separator.docx')
    expect(docx.page_properties.columns.separator).to be(false)
  end

  it 'separator is present' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/' \
                                              'columns/columns_separator/' \
                                              'with_separator.docx')
    expect(docx.page_properties.columns.separator).to be(true)
  end
end
