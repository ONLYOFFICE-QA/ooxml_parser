# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'default_document_language' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/defaults/run_properties/default_document_language.docx')
    expect(docx.styles.document_defaults.run_properties_default.run_properties.language.value).to eq('en-CA')
  end
end
