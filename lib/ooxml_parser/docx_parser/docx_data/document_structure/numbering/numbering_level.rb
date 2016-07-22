module OoxmlParser
  class NumberingLevel
    attr_accessor :ilvl, :start, :num_format, :level_text, :level_jc, :ind, :font

    def initialize(ilvl = nil, start = nil, num_fmt = nil, lvl_text = nil, lvl_jc = nil, ind = nil, font = nil)
      @ilvl = ilvl
      @start = start
      @num_format = num_fmt
      @level_text = lvl_text
      @level_jc = lvl_jc
      @ind = ind
      @font = font
    end

    def ==(other)
      if @font.nil? || other.font.nil? # Ilya Kirillov change logic of numbering font of bullets, so check should be correct
        @start.to_s == other.start.to_s && @num_format == other.num_format && @level_text == other.level_text &&
          @ind.equal_with_round(other.ind)
      else
        @start.to_s == other.start.to_s && @num_format == other.num_format && @level_text == other.level_text &&
          @ind.equal_with_round(other.ind) && @font == other.font
      end
    end
  end
end
