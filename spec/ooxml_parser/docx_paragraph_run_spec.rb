# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocxParagraphRun do
  let(:run) { described_class.new }

  describe 'equality' do
    it 'Run is equal to itself' do
      second_run = run.dup
      expect(run).to eq(second_run)
    end
  end
end
