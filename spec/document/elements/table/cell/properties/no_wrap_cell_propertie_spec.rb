# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'no_wrap_enabled' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/cell/properties/no_wrap/no_wrap_enabled.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.no_wrap).to be_truthy
  end
end
