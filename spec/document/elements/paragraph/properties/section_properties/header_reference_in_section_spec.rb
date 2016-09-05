require 'spec_helper'

describe 'My behaviour' do
  it 'header_reference_in_section' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/section_properties/header_reference/header_reference_in_section.docx')
    expect(docx.elements.first.paragraph_properties.section_properties.notes.first.elements.first.nonempty_runs.first.text).to eq('This is page header #1. ')
  end
end
