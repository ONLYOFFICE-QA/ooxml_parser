# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::SpacingValuedChild do
  let(:default_object) { described_class.new }

  describe 'SpacingValuedChild#to_ooxml_size' do
    it 'default_object to_ooxml_size raise error' do
      expect { default_object.to_ooxml_size }.to raise_error('Unknown spacing child type')
    end
  end
end
