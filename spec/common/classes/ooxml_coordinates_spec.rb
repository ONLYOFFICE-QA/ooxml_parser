# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OOXMLCoordinates do
  let(:coord) { described_class.new(1, 2) }

  it '#to_s return correct value' do
    expect(coord.to_s).to eq('(0.0 centimeter; 0.0 centimeter)')
  end
end
