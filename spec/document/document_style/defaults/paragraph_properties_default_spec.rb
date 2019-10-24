# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'default_justification' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/defaults/paragraph_properties/default_justification.docx')
    expect(docx.styles.document_defaults.paragraph_properties_default.paragraph_properties.justification).to eq(:left)
  end
end
