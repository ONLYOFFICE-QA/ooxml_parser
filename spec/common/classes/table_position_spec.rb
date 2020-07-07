# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::TablePosition do
  let(:table_position) { described_class.new }

  describe 'TablePosition#to_s' do
    it 'table_position.to_s is outputted for empty class' do
      expect(table_position.to_s).to include('Table position')
    end
  end
end
