# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OoxmlColor do
  let(:color) { described_class.new }

  describe '#==' do
    it 'OoxmlColor#== with other class' do
      expect(color).not_to eq(OoxmlParser::HeaderFooter.new)
    end
  end
end
