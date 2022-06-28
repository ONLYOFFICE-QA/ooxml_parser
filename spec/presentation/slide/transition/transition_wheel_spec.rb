# frozen_string_literal: true

require 'spec_helper'

describe 'TransitionProperties#wheel' do
  it 'tranision_wheel.pptx' do
    expect(OoxmlParser::Parser.parse('spec/presentation/slide/' \
                                     'transition/transition_wheel/' \
                                     'tranision_wheel.pptx').slides[0]
               .transition.properties
               .wheel.spokes).to eq(8)
  end
end
