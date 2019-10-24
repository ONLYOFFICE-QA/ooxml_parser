# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'anchor_lock' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/contextual_spacing/contextual_spacing_true.docx')
    expect(docx.elements.first.paragraph_properties.contextual_spacing).to be_falsey
  end
end
