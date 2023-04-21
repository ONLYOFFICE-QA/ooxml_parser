# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Underline do
  let(:underline) { described_class.new }

  describe '#to_s' do
    it '#to_s without color return correct result' do
      expect(underline.to_s).to eq('none')
    end

    it '#to_s with color return correct result' do
      underline.color = 'red'
      expect(underline.to_s).to eq('none red')
    end
  end
end
