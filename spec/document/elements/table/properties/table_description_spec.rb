# frozen_string_literal: true

require 'spec_helper'

describe 'table_caption' do
  it 'table_caption' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/properties/description/table_description.docx')
    expect(docx.elements[1].properties.description.value).to eq('SimpleTestText')
  end
end
