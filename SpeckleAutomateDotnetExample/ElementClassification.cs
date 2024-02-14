public class ElementClassification
{
    public QuantityTypes QuantityType {get;set;}
}

public enum QuantityTypes
{
    BY_VOLUME = 0,
    BY_AREA = 1,
    BY_QUANTITY = 2
}