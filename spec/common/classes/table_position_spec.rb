# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::TablePosition do
  let(:table_position) { described_class.new }

  describe 'TablePosition#to_s' do
    it 'table_position.to_s is outputted for empty class' do
      expect(table_position.to_s).to include('Table position')
    end

    it 'table_position contains all arguments in .to_s' do
      table_position = described_class.new
      table_position.left = 1
      table_position.right = 2
      table_position.top = 3
      table_position.bottom = 4
      table_position.position_x = 5
      table_position.position_y = 6
      expect(table_position.to_s).not_to be_empty
    end
  end
end
