module OoxmlParser
  class NumberingProperties
    attr_accessor :size, :font, :symbol, :start_at, :type, :ilvl, :numbering_properties

    def initialize(ilvl = 0, numbering_properties = nil)
      @ilvl = ilvl
      @numbering_properties = numbering_properties
    end

    def numbering_level_current
      @numbering_properties.ilvls.each do |current_ilvl|
        return current_ilvl if current_ilvl.ilvl == @ilvl
      end
      nil
    end
  end
end
