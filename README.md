ObservableCollection-ToExcel
Supports
		Hidden Properties
		Custom Column Header
		Custom Sheet Header
============================
Usage Exemple

private BindableCollection<Person> _Persons = new BindableCollection<Person>();

public BindableCollection<Person> Persons
{
	get
	{
		return _Persons;
	}
	set
	{
		_Persons = value;
		NotifyOfPropertyChange(() => Persons);
	}
}

List<string> HiddenProperties = new List<string>();
HiddenProperties.Add("PropertyNameToBeHidden");

Dictionary<string, string> PropertiesCustomHeader = new Dictionary<string, string>();
PropertiesCustomHeader.Add("PropertyName", "Fancy Column Header Description");

List<string> SheetCustomHeader = new List<string>();
SheetCustomHeader.Add("Description to show up in first line");
SheetCustomHeader.Add("Description to show up in second line");

Persons.ToExcel(HiddenProperties, PropertiesCustomHeader, SheetCustomHeader, true);
