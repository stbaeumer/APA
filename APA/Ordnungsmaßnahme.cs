﻿using System;

namespace APA
{
    public class Maßnahme
    {
        public int SchuelerId { get; private set; }
        public string Beschreibung { get; private set; }
        public DateTime Datum { get; private set; }
        public string Kürzel { get; private set; }
        
        public Maßnahme(int schuelerId, string beschreibung, DateTime datum, string kürzel)
        {
            SchuelerId = schuelerId;
            Beschreibung = beschreibung;
            Datum = datum;
            Kürzel = kürzel;
        }
    }
}