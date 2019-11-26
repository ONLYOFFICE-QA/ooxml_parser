# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::TableMargins do
  let(:default_margin) { described_class.new }
  let(:init_margin) { described_class.new(true, 1, 2, 3, 4) }

  describe 'TableMargins#to_s' do
    it 'default_margin to_s is outputted' do
      expect(default_margin.to_s).to include('Default: true')
    end

    it 'init_margin to_s is outputted' do
      expect(init_margin.to_s).to include('top: 1')
    end
  end
end
