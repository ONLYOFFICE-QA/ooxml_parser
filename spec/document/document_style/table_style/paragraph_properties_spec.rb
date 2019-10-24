# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'paragraph_properties_jc' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style/paragraph_properties/paragraph_properties_jc.docx')
    expect(docx.elements[1].properties.table_style.table_style_properties_list.first.paragraph_properties.justification).to eq(:center)
  end
end
