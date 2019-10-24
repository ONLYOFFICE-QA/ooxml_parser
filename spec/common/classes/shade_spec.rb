# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Shade do
  let(:shade) do
    OoxmlParser::Shade.new(value: :clear,
                           color: :auto,
                           fill: OoxmlParser::Color.new(1, 2, 3))
  end

  it 'OoxmlParser::Shade#to_s include fill color' do
    expect(shade.to_s).to include('RGB (')
  end
end
