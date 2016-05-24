require 'rspec'
require 'ooxml_parser'

describe 'My behaviour' do
  it 'Check not underlined' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/FontStyleNotUnderlined.pptx')
    font_style = pptx.slides[0].elements.last.text_body.paragraphs.first.characters.first.properties.font_style
    expect(font_style.underlined).to eq(:none)
  end

  it 'Check theme name classic' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ThemeClassic.pptx')
    expect(pptx.theme.name).to eq('Classic')
  end

  it 'Check theme name Стандартная' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/theme_standart.pptx')
    expect(pptx.theme.name).to eq('Стандартная')
  end

  it 'Check VerticalAlign Top' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/VerticalAlignTop.pptx')
    expect(pptx.slides[0].nonempty_elements.first.text_body.properties.vertical_align).to eq(:top)
  end

  it 'Check VerticalAlign bottom' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/VerticalAlignBottom.pptx')
    expect(pptx.slides[0].nonempty_elements.first.text_body.properties.vertical_align).to eq(:bottom)
  end

  it 'FontColor1' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/FontColor1.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(242, 242, 242))
  end

  it 'FontColor2' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/FontColor2.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(165, 165, 165))
  end

  it 'FontColor3' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/FontColor3.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(64, 64, 64))
  end

  it 'ImageFill' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ImageFill.pptx')
    expect(pptx.slides[0].background.fill.image.tile).not_to be_nil
  end

  it 'ChartOnSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ChartOnSlide.pptx')
    expect(pptx.slides.first.elements.first.graphic_data.first.alternate_content).not_to be_nil
  end

  it 'ChartTransformOffset' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ChartTransformOffset.pptx')
    expect(pptx.slides.last.elements.last.transform.offset.y).to eq(0)
  end

  it 'TableStyleBackgroundColor' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/TableStyleBackgroundColor.pptx')
    table = pptx.slides.first.elements.last.graphic_data.first
    expect(table.rows.first.cells[1].properties.color.color.converted_color).to eq(OoxmlParser::Color.new(68, 114, 196))
  end

  it 'TableStyleBackgroundColor Empty' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/TableStyleBackgroundColor.pptx')
    table = pptx.slides.first.elements.last.graphic_data.first
    expect(table.rows.first.cells.first.properties.color).to be_nil
  end

  it 'Character Spacing 1cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/CharacterSpacing1cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(1)
  end

  it 'Character Spacing 5cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/CharacterSpacing5cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(5)
  end

  it 'Character Spacing 0.1cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/CharacterSpacing0.1cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(0.1)
  end

  it 'Check image size' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ImageSize.pptx')
    transform = pptx.slides.first.elements.last.properties.transform
    expect(transform.extents.x).to eq(13)
    expect(transform.extents.y).to eq(5)
  end

  it 'Check image shift' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ImageShift.pptx')
    transform = pptx.slides.first.elements.last.properties.transform
    expect(transform.offset.x).to eq(pptx.slide_size.width)
    expect(transform.offset.y).to eq(pptx.slide_size.height)
  end

  it 'HyperlinkToSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/HyperlinkToSlide.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:previous_slide)
  end

  it 'HyperlinkToSpecificSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/HyperlinkToSpecificSlide.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:slide)
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.link_to).to eq(1)
  end

  it 'LinkInTextArea' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/LinkInTextArea.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:external_link)
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.link_to).to eq('http://www.yandex.ru/')
  end

  it 'Align left' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/Autosave.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.align).to eq(:left)
  end

  it 'Spacing' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/Spacing.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :multiple))
  end

  it 'DefaultSpace' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/DefaultSpace.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 1, :multiple))
  end

  it 'ShapeGrouping' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ShapeGrouping.pptx')
    expect(pptx.slides[0].elements[0]).to be_a_kind_of OoxmlParser::ShapesGrouping
  end

  it 'Shape Naming' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ShapeNaming.pptx')
    expect(pptx.slides[0].elements[-1].shape_properties.preset).to eq(:rect)
  end

  it 'PresentationComment' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/PresentationComment.pptx')
    expect(pptx.comments.first.text).to eq('Is it true?')
  end

  it 'SlidePattern.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SlidePattern.pptx')
    expect(pptx.slides[0].background.fill.pattern
               .background_color.value).to eq(OoxmlParser::Color.new(231, 230, 230))
  end

  it 'SlideGradient.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SlideGradient.pptx')
    gradient_stops = pptx.slides[0].background.fill.color.gradient_stops
    expect(gradient_stops[0].color.value).to eq(OoxmlParser::Color.new(237, 125, 49))
    expect(gradient_stops[0].position).to eq(10)
    expect(gradient_stops[1].position).to eq(70)
  end

  it 'GroupShapes.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/GroupShapes.pptx')
    expect(pptx.slides.first.elements.first.elements.first.shape_properties.preset).to eq(:ellipse)
    expect(pptx.slides.first.elements.first.elements.last.shape_properties.preset).to eq(:mathPlus)
  end

  describe 'Spacing Rule' do
    it 'SpacingLineRuleExact.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SpacingLineRuleExact.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(1.84, 1.92, 0.8, :exact))
    end

    it 'SpacingLineRuleExact0.35.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SpacingLineRuleExact0.35.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 0.35, :exact))
    end

    it 'SpacingLineRuleExact1.76.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SpacingLineRuleExact1.76.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 1.76, :exact))
    end

    it 'SpacingLineRuleMultiple.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/SpacingLineRuleMultiple.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(1.84, 1.92, 1, :multiple))
    end
  end

  it 'ColumnCharts.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ColumnCharts.pptx')
    elements = pptx.slides.first.elements
    expect(elements[0].graphic_data.first.type).to eq(:column)
    expect(elements[0].graphic_data.first.grouping).to eq(:clustered)
    expect(elements[1].graphic_data.first.type).to eq(:column)
    expect(elements[1].graphic_data.first.grouping).to eq(:stacked)
    expect(elements[2].graphic_data.first.type).to eq(:column)
    expect(elements[2].graphic_data.first.grouping).to eq(:percentStacked)
  end

  it 'PresentationPictureExists.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/PresentationPictureExists.pptx')
    expect(pptx.slides.first.elements.size).to be >= 3
    expect(pptx.slides.first.elements.last).to be_a_kind_of OoxmlParser::Picture
    expect(File.size?(pptx.slides.first.elements.last.image.path)).to be > 0
  end

  it 'GroupingChart.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/GroupingChart.pptx')
    elements = pptx.slides.first.elements
    expect(elements.first).to be_a_kind_of OoxmlParser::ShapesGrouping
    expect(elements[0].elements.size).to eq(2)
  end

  it 'ChartBorders.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ChartBorders.pptx')
    expect(pptx.slides[0].elements.first.graphic_data.first.shape_properties.line.width).to eq(0)
    expect(pptx.slides[1].elements.first.graphic_data.first.shape_properties.line.color).to be_nil
    expect(pptx.slides[2].elements.first.graphic_data.first.shape_properties.line.width).to eq(0.5)
    expect(pptx.slides[2].elements.first.graphic_data.first.shape_properties.line.color).not_to be_nil
  end

  it 'TableColor.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/TableColor.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    expect(table.properties.style.first_row.cell_style.fill.color).to eq(OoxmlParser::Color.new(91, 155, 213))
  end

  it 'TableBorderStyle.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/TableBorderStyle.pptx')
    table = pptx.slides.first.graphic_frames.first.graphic_data.first
    table.rows.each do |current_row|
      current_row.cells.each do |current_cell|
        current_cell.properties.borders.each_side do |current_side|
          expect(current_side.fill.color.converted_color).to eq(OoxmlParser::Color.new(165, 165, 165))
          expect(current_side.width).to eq(4.5)
        end
        expect(current_cell.properties.color.color.converted_color).to eq(OoxmlParser::Color.new(68, 114, 196))
      end
    end
  end

  it 'Numbering.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/Numbering.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0]
               .properties.numbering.type).to eq(:alphaUcPeriod)
  end

  it 'Bullet.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/Bullet.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0]
               .properties.numbering.symbol).to eq('•')
  end

  describe 'Copy style in table' do
    it 'copy_style_in_table_with_bold.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/copy_style_in_table_with_bold.pptx')
      expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
    end

    it 'copy_style_in_table_without_bold.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/copy_style_in_table_without_bold.pptx')
      expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_falsey
    end

    it 'copy_style_in_table_with_bold_2_cell_with_bold_text_too.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/copy_style_in_table_with_bold_2_cell_with_bold_text_too.pptx')
      expect(pptx.slides[0].elements.last.graphic_data.first.rows.first.cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
      expect(pptx.slides[0].elements.last.graphic_data.first.rows[1].cells.first.text_body.paragraphs.first.runs.first.properties.font_style.bold).to be_truthy
    end
  end

  it 'TableLook.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/TableLook.pptx')
    table_look = pptx.slides.first.elements.last.graphic_data[0].properties.table_look

    expect(table_look.first_column).to eq(false)
    expect(table_look.first_row).to eq(true)
    expect(table_look.last_column).to eq(false)
    expect(table_look.last_row).to eq(false)
    expect(table_look.banding_row).to eq(false)
    expect(table_look.banding_column).to eq(false)
  end

  it 'ChartStrokeNoLine.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ChartStrokeNoLine.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.shape_properties.line.width).to eq(0)
  end

  it 'ChartStroke0.5.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/ChartStroke0.5.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.shape_properties.line.width).to eq(0.5)
  end

  describe 'chart' do
    describe 'title' do
      it 'ChartTitleBold.pptx' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/chart/title/ChartTitleBold.pptx')
        drawing = pptx.slides[0].elements.last
        expect(drawing.graphic_data.first.title.elements.first.runs.first.properties.font_style.bold).to eq(true)
      end

      it 'ChartTitleWithoutBold.pptx' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/chart/title/ChartTitleWithoutBold.pptx')
        drawing = pptx.slides[0].elements.last
        expect(drawing.graphic_data.first.title.elements.first.runs.first.properties.font_style.bold).to eq(false)
      end
    end
  end

  describe 'editor_specific_documents' do
    describe 'libreoffice' do
      it 'simple_text' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/editor_specific_documents/libreoffice/simple_text.pptx')
        drawing = pptx.slides[0].elements.last
        expect(drawing.text_body.paragraphs.first.runs.first.text).to eq('This is a test')
      end
    end

    describe 'ms_office_2013' do
      it 'simple_text' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/editor_specific_documents/ms_office_2013/simple_text.pptx')
        drawing = pptx.slides[0].elements.last
        expect(drawing.text_body.paragraphs.first.runs.first.text).to eq('This is a test')
      end
    end
  end

  it 'FontStyleUnderlineNone.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/FontStyleUnderlineNone.pptx')
    expect(pptx.slides[0].elements.last.text_body.paragraphs.first.characters.first.properties.font_style).to eq(OoxmlParser::FontStyle.new)
  end

  describe 'Parsing Theme' do
    it 'NoThemeName' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/theme/NoThemeName.pptx')
      expect(pptx.theme.name).to eq('')
    end

    it 'ClassicThemeName' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/theme/ClassicThemeName.pptx')
      expect(pptx.theme.name).to eq('Classic')
    end
  end

  describe 'Parsing Table' do
    it 'TableCellVerticalAlignTop.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/TableCellVerticalAlignTop.pptx')
      first_row = pptx.slides.first.elements.last.graphic_data.first.rows.first
      expect(first_row.cells.first.properties.anchor).to eq(:top)
    end

    it 'TableCellVerticalAlignBottom.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/TableCellVerticalAlignBottom.pptx')
      first_row = pptx.slides.first.elements.last.graphic_data.first.rows.first
      expect(first_row.cells.first.properties.anchor).to eq(:bottom)
    end

    it 'table_background_color.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/table/table_background_color.pptx')
      table = pptx.slides.first.graphic_frames.first.graphic_data.first
      expect(OoxmlParser::Color.to_color(table.rows.first.cells.first.properties.color)).to eq(OoxmlParser::Color.new(70, 93, 73))
    end
  end

  describe 'Transition' do
    it 'transition_direction.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/transition/transition_direction.pptx')
      expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:right_down)
    end

    it 'transition_direction_down.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/transition/transition_direction_down.pptx')
      expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:down)
    end

    it 'transition_direction_up.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/transition/transition_direction_up.pptx')
      expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:up)
    end

    it 'transition_direction_in.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/transition/transition_direction_in.pptx')
      expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:in)
    end
  end

  describe 'shape' do
    describe 'text' do
      describe 'paragraph' do
        describe 'properties' do
          it 'tranistion_sound_action' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/properties/tranistion_sound_action.pptx')
            expect(File.exist?(pptx.slides.first.transition.sound_action.start_sound.path)).to be_truthy
          end

          it 'transition_direction_nil_1' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/properties/transition_direction_nil_1.pptx')
            expect(pptx.slides.first.transition.properties.direction).to be_nil
          end

          it 'transition_direction_nil_2' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/properties/transition_direction_nil_2.pptx')
            expect(pptx.slides[1].transition.properties.direction).to be_nil
          end

          it 'transition_no_orientation' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/properties/transition_no_orientation.pptx')
            expect(pptx.slides.first.transition.properties.orientation).to be_nil
          end
        end

        describe 'run' do
          it 'strikeout.pptx' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/run/strikeout.pptx')
            expect(pptx.slides.first.elements.first.text_body.paragraphs.first.runs.first.properties.font_style.strike).to eq(:single)
          end

          it 'several_paragraphs_font_style.pptx' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/run/several_paragraphs_font_style.pptx')
            expect(pptx.slides.first.elements.first.text_body.paragraphs[0].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(false, false, OoxmlParser::Underline.new(:single), :none))
            expect(pptx.slides.first.elements.first.text_body.paragraphs[1].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(true))
            expect(pptx.slides.first.elements.first.text_body.paragraphs[2].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(false, false, OoxmlParser::Underline.new(:single), :none))
            expect(pptx.slides.first.elements.first.text_body.paragraphs[3].runs.first.properties.font_style).to eq(OoxmlParser::FontStyle.new(true))
          end

          describe 'defaults' do
            it 'DefaultFontName' do
              pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/run/defaults/DefaultFontName.pptx')
              expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_name).to eq('Arial')
            end

            it 'DefaultFontSize.pptx' do
              pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/shape/text/paragraph/run/defaults/DefaultFontSize.pptx')
              expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_size).to eq(18)
            end
          end
        end
      end
    end
  end

  describe 'slide' do
    describe 'timing' do
      describe 'time_node' do
        describe 'condition' do
          it 'condition_no_event' do
            pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/slide/timing/time_node/condition/condition_no_event.pptx')
            expect(pptx.slides[1].timing.time_node_list.first.common_time_node.children.first
                       .common_time_node.children.first.common_time_node.children.first.common_time_node.start_conditions.first.event).to be_nil
          end
        end
      end
    end
  end
end
