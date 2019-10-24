# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ImageSize' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/image_size.xlsx')
    expect(xlsx.worksheets.first.drawings).not_to be_empty
    expect(xlsx.worksheets.first.drawings.first.picture.path_to_image.file_reference.content.length).to be > 0
  end

  it 'picture_with_non_visual' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/picture_with_non_visual.xlsx')
    expect(xlsx.worksheets.first.drawings.first.picture.non_visual_properties).to be_a(OoxmlParser::NonVisualShapeProperties)
  end

  it 'Incorrect image resource' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/incorrect_image_resource.xlsx') }.to raise_error(LoadError)
  end

  it 'image_incorrect_link.xlsx' do
    expect { OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/image_incorrect_link.xlsx') }
      .to output("Cant find path to media file by id: rId1\n").to_stderr
  end
end
