# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'parser will not crash if there is no default run style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/defaults/run_properties/no_default_run_properties.docx')
    expect(docx).to be_with_data
  end

  it 'parser will not crash if there is no run properties in document_defaults' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/defaults/run_properties/no_run_properties_document_defaults.docx')
    expect(docx).to be_with_data
  end

  it 'default_document_language' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/defaults/run_properties/default_document_language.docx')
    expect(docx.styles.document_defaults.run_properties_default.run_properties.language.value).to eq('en-CA')
  end
end
