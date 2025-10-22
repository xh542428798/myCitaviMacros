using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Citations;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Collections;

namespace SwissAcademic.Citavi.Comparers
{
	public class CustomCitationComparer
     :
     ICustomCitationComparerMacro
	{
		public int Compare(Citation x, Citation y)
	  	{
		    var defaultCitationComparer = CitationComparer.AuthorYearTitleAscending;
			
			var xReference = x.Reference;
			var yReference = y.Reference;

			string xCriterionOne;
			string xCriterionTwo;
			string xCriterionThree;
			string xCriterionFour;
			string xCriterionFive;

			string yCriterionOne;
			string yCriterionTwo;
			string yCriterionThree;
			string yCriterionFour;
			string yCriterionFive;

			GetCriteria(xReference, out xCriterionOne, out xCriterionTwo, out xCriterionThree, out xCriterionFour, out xCriterionFive);
			GetCriteria(yReference, out yCriterionOne, out yCriterionTwo, out yCriterionThree, out yCriterionFour, out yCriterionFive);

			if (xCriterionOne != yCriterionOne)
			{
				return xCriterionOne.CompareTo(yCriterionOne);
			}
			else if (xCriterionTwo != yCriterionTwo)
			{
				return xCriterionTwo.CompareTo(yCriterionTwo);
			}
			else if (xCriterionThree != yCriterionThree)
			{
				return xCriterionThree.CompareTo(yCriterionThree);
			}
			else if (xCriterionFour != yCriterionFour)
			{
				return -xCriterionFour.CompareTo(yCriterionFour);
			}
			else if (xCriterionFive != yCriterionFive)
			{
				return xCriterionFive.CompareTo(yCriterionFive);
			}
			else
			{
				return 0;
			}
	  	}
		private void GetCriteria(Reference reference, out string criterionOne, out string criterionTwo, out string criterionThree, out string criterionFour, out string criterionFive)
		{
			criterionOne = GetCriterionOne(reference);
			criterionTwo = GetCriterionTwo(reference);
			criterionThree = GetCriterionThree(reference);;
			criterionFour = GetCriterionFour(reference);
			criterionFive = GetCriterionFive(reference);
			return;
		}
		private string GetCriterionOne(Reference reference)
		{
			string criterion = "0";
			return criterion;
		}
		private string GetCriterionTwo(Reference reference)
		{
			string criterion = "0";

			if (reference.ReferenceType == ReferenceType.StatuteOrRegulation || reference.ReferenceType == ReferenceType.CourtDecision)
			{
				IList<Person> courts = reference.Organizations as IList<Person>;
				if (courts == null) return "0";
				
				Person court = courts.FirstOrDefault();
				if (court == null) return "0";
				
				string jurisdiction = court.LastNameForSorting;
				if (JurisdictionIsGermany(jurisdiction))
				{
					criterion = "Germany";
				}
				else
				{
					criterion = jurisdiction;
				}
			}
			else			
			{
				criterion = reference.AuthorsOrEditorsOrOrganizations.ToStringSafe();
			}
			return criterion;
		}
		private string GetCriterionThree(Reference reference)
		{
			string criterion = "0";
			if (reference.ReferenceType == ReferenceType.CourtDecision) criterion = GetCriterionThreeCourtDecision(reference);
			else criterion = reference.Title;
			return criterion;
		}
		private string GetCriterionThreeCourtDecision(Reference reference)
		{
			if (reference.ReferenceType != ReferenceType.CourtDecision) return "0";
			string criterion = "0";

			IList<Person> courts = reference.Organizations as IList<Person>;
			if (courts == null) return "0";

			Person court = courts.FirstOrDefault();
			if (court == null) return "0";

			string jurisdiction = court.LastNameForSorting;
			if (string.IsNullOrEmpty(jurisdiction)) return "0";

			string theCourt = string.Empty;

			if (JurisdictionIsFrance(jurisdiction) || JurisdictionIsGermany(jurisdiction) || JurisdictionIsNetherlands(jurisdiction))
			{
				if (!string.IsNullOrEmpty(court.Abbreviation))
				{
					if (!string.IsNullOrEmpty(theCourt)) theCourt += ", ";
					theCourt += court.Abbreviation;
				}
				else if (!string.IsNullOrEmpty(court.LastName))
				{
					if (!string.IsNullOrEmpty(theCourt)) theCourt += ", ";
					theCourt += court.LastName;
				}
				criterion = GetCourtRanking(theCourt);
			}
			else if (!String.IsNullOrEmpty(reference.Title))
			{
				criterion = reference.Title;
				if (criterion.StartsWith("Re ")) criterion = criterion.Remove(0, 3);
				if (criterion.StartsWith("R v.")) criterion = criterion.Remove(0, 4);
			}
			return criterion;
		}
		private string GetCriterionFour(Reference reference)
		{
			string criterion = "0";
			if (!String.IsNullOrEmpty(reference.Date.ToStringSafe())) criterion = reference.Date.ToStringSafe();
			else criterion = reference.YearResolved.ToStringSafe();
			return criterion;
		}
		private string GetCriterionFive(Reference reference)
		{
			string criterion = "0";
			if (reference.ReferenceType == ReferenceType.CourtDecision)
			{
				criterion = GetPeriodicalRanking(reference);
			}
			return criterion;
		}
		private string GetCourtRanking(string court)
		{
			if (string.IsNullOrEmpty(court)) return "zz";

			// Oberste Bundesgerichte

			else if (court.Equals("Cons. d'État")) return "ca";
			else if (court.Equals("Cass.")) return "cba";
			else if (court.Equals("Cass. ass. plén.")) return "cba";
			else if (court.Equals("Cass. 1ère")) return "cbb";
			else if (court.Equals("Cass. 2ème")) return "cbb";
			else if (court.Equals("Comm.")) return "cbb";			

			else if (court.Equals("BVerfG")) return "ca";
			else if (court.Equals("RG")) return "cab";
			else if (court.Equals("BGH")) return "cb";
			else if (court.Equals("BVerwG")) return "cc";
			else if (court.Equals("BAG")) return "cd";
			else if (court.Equals("BFG")) return "ce";
			else if (court.Equals("BSG")) return "cf";

			else if (court.Equals("HR")) return "ca";
			else if (court.Equals("PHR")) return "cab";

			// Bundesobergerichte

			else if (court.Equals("BPatg")) return "da";

			// Landesverfassungsgerichte

			else if (court.Contains("LVerfG")) return "dn";

			// Bayerisches Oberstes Landesgericht

			else if (court.Contains("Bay. Ob. LG")) return "do";

			// Landesobergerichte

			else if (court.Contains("CA")) return "ea";

			else if (court.Contains("OLG") || court.Contains("KG")) return "ea";
			else if
				(
				court.Contains("OVG")
				|| court.Contains("VGH")
				|| court.Contains("C.A.A.")
				)
				return "eb";
			else if (court.Contains("LArbG")) return "ec";
			else if (court.Contains("LFG")) return "ed";
			else if (court.Contains("LSG")) return "ee";

			else if (court.Contains("Rechtbank")) return "ea";

			// Landesgerichte mit allgemeiner Zustaendigkeit

			else if (court.Contains("LG")) return "fa";
			else if
				(
					court.Contains("VG")
					|| court.Contains("T.A.")
				) return "fb";
			else if (court.Contains("ArbG")) return "fc";
			else if (court.Contains("FG")) return "fd";
			else if (court.Contains("LSG")) return "fe";

			// Amtsgerichte

			else if (court.Contains("AG")) return "ge";

			return "zz";
		}
		private string GetPeriodicalRanking(Reference reference)
		{
			string periodicalString = string.Empty;

			if (reference.Periodical == null) return "0";

			if (!string.IsNullOrEmpty(reference.Periodical.StandardAbbreviation))
			{
				periodicalString = reference.Periodical.StandardAbbreviation;
			}
			else if (!string.IsNullOrEmpty(reference.Periodical.FullName))
			{
				periodicalString = reference.Periodical.FullName;
			}

			if (string.IsNullOrEmpty(periodicalString)) return "zz";

			else if (reference.Periodical.Notes.ToStringSafe().Contains("neutral")) return "aa";

			// Here we leave room for customizations in case a publisher wants you to cite one reporter over another

			else if (reference.Periodical.Notes.ToStringSafe().Contains("official")) return "na";
			else if (periodicalString.Contains("W.L.R.")) return "ob"; // WLR should always come after the more specialized ICLR reports
			else if (reference.Periodical.Notes.ToStringSafe().Contains("semi")) return "oa";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("nominate")) return "pa";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("national")) return "qa";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("regional")) return "ra";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("local")) return "sa";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("topical")) return "ta";
			else if (reference.Periodical.Notes.ToStringSafe().Contains("electronic")) return "ua";

			else return "zz";
		}
		private void EliminateParallelReporters(Citation x, Citation y)
		{
			var xBibliographyCitation = x as BibliographyCitation;
			var yBibliographyCitation = y as BibliographyCitation;

			var xReference = x.Reference;
			var yReference = y.Reference;

			if (xReference.Periodical == null) return;
			if (yReference.Periodical == null) return;

			var xIsCourtDecision = xReference.ReferenceType == ReferenceType.CourtDecision;
			var yIsCourtDecision = yReference.ReferenceType == ReferenceType.CourtDecision;


			// We check whether y is a parallel reporter of x
			if (xIsCourtDecision &&
				yIsCourtDecision &&
				xReference.Organizations.SequenceEqual(yReference.Organizations) &&
				((xReference.Date == yReference.Date &&
				  xReference.Title == yReference.Title) ||
				 (xReference.Date == yReference.Date &&
				  xReference.SpecificField2 == yReference.SpecificField2)))
			{
				string xPeriodicalRanking = GetPeriodicalRanking(xReference);
				string yPeriodicalRanking = GetPeriodicalRanking(yReference);

				// If both periodicals are ranked differently, we kick out the lower ranked one
				if (xPeriodicalRanking != yPeriodicalRanking)
				{
					if (xPeriodicalRanking.CompareTo(yPeriodicalRanking) < 0) yBibliographyCitation.NoBib = true;
					if (xPeriodicalRanking.CompareTo(yPeriodicalRanking) > 0) xBibliographyCitation.NoBib = true;
				}
				//If two periodicals are of equal ranking, we use the title of the periodical to break the tie
				else
				{
					string xPeriodicalString = string.Empty;
					string yPeriodicalString = string.Empty;

					if (!string.IsNullOrEmpty(xReference.Periodical.StandardAbbreviation))
					{
						xPeriodicalString += xReference.Periodical.StandardAbbreviation;
					}
					else if (!string.IsNullOrEmpty(xReference.Periodical.FullName))
					{
						xPeriodicalString += xReference.Periodical.FullName;
					}
					if (!string.IsNullOrEmpty(yReference.Periodical.StandardAbbreviation))
					{
						yPeriodicalString += yReference.Periodical.StandardAbbreviation;
					}
					else if (!string.IsNullOrEmpty(yReference.Periodical.FullName))
					{
						yPeriodicalString += yReference.Periodical.FullName;
					}

					if (xPeriodicalString.CompareTo(yPeriodicalString) < 0) yBibliographyCitation.NoBib = true;
					if (xPeriodicalString.CompareTo(yPeriodicalString) > 0) xBibliographyCitation.NoBib = true;
				}
			}
		}
		//////////////////////////////////////////////////////
		//				JURISDICTION TESTS					//
		//////////////////////////////////////////////////////
		static bool JurisdictionIsEurope(string jurisdiction)
		{
			List<string> jurisdictionsEU = new List<string>
			{
				"EU"
			};
			return jurisdictionsEU.Any(jurisdiction.Equals);
		}
		static bool JurisdictionIsFrance(string jurisdiction)
		{
			List<string> jurisdictionsFrance = new List<string>
			{
				"F",
			};
			return jurisdictionsFrance.Any(jurisdiction.Equals);
		}
		static bool JurisdictionIsGermany(string jurisdiction)
		{
			List<string> jurisdictionsGermany = new List<string>
			{
				"D",
				"Br. Z.",
				"BB", "BE", "BW", "BY", "HB", "HE", "HH", "MV", "NI", "NW", "RP", "SH", "SN", "SL", "ST", "TH",
			};
			return jurisdictionsGermany.Any(jurisdiction.Equals);
		}
		static bool JurisdictionIsNetherlands(string jurisdiction)
		{
			List<string> jurisdictionsNetherlands = new List<string>
			{
				"NL",
			};
			return jurisdictionsNetherlands.Any(jurisdiction.Equals);
		}
		//////////////////////////////////////////////////////
		//			END OF JURISDICTION TESTS			//
		//////////////////////////////////////////////////////
	}
}