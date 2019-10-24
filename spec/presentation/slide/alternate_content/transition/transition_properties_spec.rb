# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'empty_transition_properties.pptx' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/alternate_content/transition/properties/empty_transition_properties.pptx')
    expect(pptx.slides.first.alternate_content.transition.properties).to be_nil
  end
end
