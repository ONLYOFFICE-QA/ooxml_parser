# frozen_string_literal: true

require 'spec_helper'

describe 'Table Style' do
  it 'TableStyle root object 1' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_style/table_style_root_object_1.docx')
    expect(docx).to be_with_data
  end

  it 'TableStyle root object 2' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_style/table_style_root_object_2.docx')
    expect(docx).to be_with_data
  end
end
