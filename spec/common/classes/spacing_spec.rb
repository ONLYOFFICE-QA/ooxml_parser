# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Spacing do
  let(:spacing) do
    OoxmlParser::Spacing.new(1.05, 2.10, 3.15, :at_least)
  end

  describe 'Equality' do
    it 'Spacing is not equal nil' do
      expect(spacing).not_to be_nil
    end

    it 'Spacing is not equal with some other spacing' do
      expect(spacing).not_to eq(OoxmlParser::Spacing.new)
    end
  end

  it 'Spacing after copy is eqaul for itself' do
    expect(spacing).to eq(spacing.copy)
  end

  it 'Spacing is correct after round' do
    expect(spacing.round(1)).to eq(OoxmlParser::Spacing.new(1.1, 2.1, 3.2, :at_least))
  end
end
