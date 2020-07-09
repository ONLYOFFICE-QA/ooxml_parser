# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OoxmlSize do
  let(:size) { described_class.new(1, :centimeter) }

  describe '#to_unit' do
    it 'OoxmlSize to other unit is same to base unit' do
      expect(size.to_unit('foo')).to eq(size.to_base_unit)
    end

    it 'OoxmlSize to one_240th_cm' do
      expect(size.to_unit(:one_240th_cm))
        .to eq(described_class.new(240, :one_240th_cm))
    end
  end

  describe '#to_s' do
    it 'OoxmlSize to_s output is correct' do
      expect(size.to_s).to eq('1.0 centimeter')
    end
  end
end
