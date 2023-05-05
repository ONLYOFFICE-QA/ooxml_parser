# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ClientData do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/client_data/client_data.xlsx')

  it 'Has locks_with_sheet' do
    expect(xlsx.worksheets[0].drawings[0].client_data.locks_with_sheet).to be_falsey
  end
end
