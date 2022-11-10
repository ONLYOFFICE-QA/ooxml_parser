# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'anchor_lock' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/contextual_spacing/contextual_spacing_true.docx')
    expect(docx.elements.first.paragraph_properties.contextual_spacing).to be_falsey
  end

  it 'different contextual spacing' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/contextual_spacing/different_contextual_spacing.docx')
    expect(docx.elements[0].paragraph_properties.contextual_spacing).to be_falsey
    expect(docx.elements[1].paragraph_properties.contextual_spacing).to be_falsey
    expect(docx.elements[2].paragraph_properties.contextual_spacing).to be_falsey
    expect(docx.elements[3].paragraph_properties.contextual_spacing).to be_truthy
    expect(docx.elements[4].paragraph_properties.contextual_spacing).to be_truthy
  end
end
