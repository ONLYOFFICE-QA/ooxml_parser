# frozen_string_literal: true

require 'spec_helper'

describe 'keep_next' do
  it 'keep_next' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/properties/keep_next/keep_next.docx')
    expect(docx.elements[1].paragraph_properties.keep_next).to be_truthy
  end
end
