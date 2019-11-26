# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::BordersProperties do
  let(:default_property) { described_class.new }
  let(:init_property) { described_class.new(:auto, 1) }

  describe 'BordersProperties#to_s' do
    it 'default_property to_s is empty' do
      expect(default_property.to_s).to be_empty
    end

    it 'init_property to_s with data' do
      expect(init_property.to_s).to include('auto')
    end
  end

  it 'BordersProperties#copy' do
    copy = init_property.copy
    expect(copy).to eq(init_property)
  end

  describe 'BordersProperties#visible?' do
    it 'default_property is not visible' do
      expect(default_property).not_to be_visible
    end

    it 'init_property is visible' do
      expect(init_property).to be_visible
    end
  end
end
