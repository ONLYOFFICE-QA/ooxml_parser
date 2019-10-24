# frozen_string_literal: true

require 'spec_helper'

describe 'table_caption' do
  it 'table_caption' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/properties/caption/table_caption.docx')
    expect(docx.elements[1].properties.caption.value).to eq('Simple Test Text')
  end
end
